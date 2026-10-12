/**
 * استيراد "خطة 0" (الرصيد الابتدائي) من ملف Excel
 *
 * الاستخدام (من مجلد الـ backend):
 *   node scripts/importInitialStats.js ./data.xlsx            # معاينة فقط (لا يكتب شيئًا)
 *   node scripts/importInitialStats.js ./data.xlsx --apply    # كتابة في قاعدة البيانات
 *   node scripts/importInitialStats.js ./data.xlsx --apply --exact-only
 *
 * الملف: العمود الأول = اسم المشرف، العمود الثاني = العدد (الصف الأول عناوين).
 */

const path = require("path");
const xlsx = require("xlsx");

const pool = require("../models/db");

// -----------------------------------------------------
// تطبيع الأسماء العربية للمقارنة
// -----------------------------------------------------

function normalizeName(value) {
  return String(value ?? "")
    .replace(/[\u064B-\u065F\u0670\u0640]/g, "") // تشكيل + تطويل
    .replace(/[أإآ]/g, "ا")
    .replace(/ى/g, "ي")
    .replace(/ة/g, "ه")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function tokensOf(value) {
  return normalizeName(value).split(" ").filter(Boolean);
}

// -----------------------------------------------------
// قراءة الإكسل
// -----------------------------------------------------

function readExcelRows(filePath) {
  const workbook = xlsx.readFile(filePath);

  const sheet = workbook.Sheets[workbook.SheetNames[0]];

  const raw = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: null });

  return raw
    .slice(1) // تخطي صف العناوين
    .map((row, index) => ({
      excelRow: index + 2,
      name: String(row[0] ?? "").trim(),
      count: Number(row[1]),
    }))
    .filter((row) => row.name);
}

// -----------------------------------------------------
// مطابقة اسم الإكسل مع المشرفين
// -----------------------------------------------------

function matchSupervisor(excelName, supervisors) {
  const normalized = normalizeName(excelName);

  const exact = supervisors.filter(
    (s) => normalizeName(s.name) === normalized,
  );

  if (exact.length === 1) {
    return { type: "exact", supervisor: exact[0] };
  }

  if (exact.length > 1) {
    return { type: "ambiguous", candidates: exact };
  }

  // مطابقة تقريبية: كل كلمات اسم الإكسل موجودة في اسم المشرف (أو العكس)
  const excelTokens = tokensOf(excelName);

  const fuzzy = supervisors.filter((s) => {
    const supTokens = tokensOf(s.name);

    const excelInSup = excelTokens.every((t) => supTokens.includes(t));
    const supInExcel = supTokens.every((t) => excelTokens.includes(t));

    return excelInSup || supInExcel;
  });

  if (fuzzy.length === 1) {
    return { type: "fuzzy", supervisor: fuzzy[0] };
  }

  if (fuzzy.length > 1) {
    return { type: "ambiguous", candidates: fuzzy };
  }

  return { type: "none" };
}

// -----------------------------------------------------
// Main
// -----------------------------------------------------

async function main() {
  const args = process.argv.slice(2);

  const filePath = args.find((a) => !a.startsWith("--"));

  const apply = args.includes("--apply");
  const exactOnly = args.includes("--exact-only");

  if (!filePath) {
    console.error("Usage: node scripts/importInitialStats.js <file.xlsx> [--apply] [--exact-only]");
    process.exit(1);
  }

  const rows = readExcelRows(path.resolve(filePath));

  const { rows: supervisors } = await pool.query(
    `SELECT id, name FROM supervisors ORDER BY id`,
  );

  console.log(`📄 Excel rows: ${rows.length} | 👥 Supervisors in DB: ${supervisors.length}`);
  console.log(apply ? "✍️  MODE: APPLY" : "👀 MODE: DRY-RUN (لن يتم الكتابة)");

  const toWrite = new Map(); // supervisorId -> { count, excelName }
  const problems = [];

  for (const row of rows) {
    if (!Number.isInteger(row.count) || row.count < 0) {
      problems.push(`صف ${row.excelRow} «${row.name}»: قيمة غير صالحة (${row.count})`);
      continue;
    }

    const match = matchSupervisor(row.name, supervisors);

    if (match.type === "none") {
      problems.push(`صف ${row.excelRow} «${row.name}»: لا يوجد مشرف مطابق`);
      continue;
    }

    if (match.type === "ambiguous") {
      problems.push(
        `صف ${row.excelRow} «${row.name}»: أكثر من مشرف محتمل → ${match.candidates
          .map((c) => `${c.name} (#${c.id})`)
          .join(" | ")}`,
      );
      continue;
    }

    if (match.type === "fuzzy" && exactOnly) {
      problems.push(
        `صف ${row.excelRow} «${row.name}»: مطابقة تقريبية مع «${match.supervisor.name}» (تم التخطي بسبب --exact-only)`,
      );
      continue;
    }

    const id = match.supervisor.id;

    if (toWrite.has(id)) {
      problems.push(
        `صف ${row.excelRow} «${row.name}»: نفس المشرف «${match.supervisor.name}» مطابق أيضًا لـ «${toWrite.get(id).excelName}» — تم التخطي، راجعيهما يدويًا`,
      );
      continue;
    }

    toWrite.set(id, { count: row.count, excelName: row.name });

    console.log(
      `${match.type === "exact" ? "✅" : "🟡"} ${row.name}  →  ${match.supervisor.name} (#${id})  =  ${row.count}  [${match.type}]`,
    );
  }

  console.log("\n==================== مشاكل تحتاج مراجعة ====================");
  problems.length ? problems.forEach((p) => console.log("⚠️ ", p)) : console.log("لا يوجد");

  console.log(`\nجاهز للكتابة: ${toWrite.size} / ${rows.length}`);

  if (!apply) {
    console.log("\nلتنفيذ الكتابة أعيدي التشغيل مع --apply");
    return;
  }

  const client = await pool.connect();

  try {
    await client.query("BEGIN");

    for (const [supervisorId, { count }] of toWrite) {
      await client.query(
        `
        INSERT INTO supervisor_initial_stats (supervisor_id, initial_count, updated_at)
        VALUES ($1, $2, NOW())
        ON CONFLICT (supervisor_id)
        DO UPDATE SET initial_count = EXCLUDED.initial_count, updated_at = NOW()
        `,
        [supervisorId, count],
      );
    }

    await client.query("COMMIT");

    console.log(`\n💾 تم حفظ ${toWrite.size} رصيد ابتدائي.`);
  } catch (error) {
    await client.query("ROLLBACK");
    throw error;
  } finally {
    client.release();
  }
}

main()
  .catch((error) => {
    console.error("❌ Import failed:", error);
    process.exitCode = 1;
  })
  .finally(() => pool.end());