const ExcelJS = require("exceljs");
const pool = require("../models/db");
const { v4: uuidv4 } = require("uuid");

// =====================================================
// Helpers
// =====================================================

function excelDateToJSDate(value) {
  if (value instanceof Date) {
    return value;
  }

  if (typeof value === "number") {
    const utcDays = Math.floor(value - 25569);
    const utcValue = utcDays * 86400;
    return new Date(utcValue * 1000);
  }

  if (typeof value === "string") {
    const parsed = new Date(value);

    if (!Number.isNaN(parsed.getTime())) {
      return parsed;
    }
  }

  return null;
}

// =====================================================
// Normalize Excel Time
// =====================================================

function normalizeTime(value) {
  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {
    return null;
  }

  // ExcelJS may return a Date object
  if (value instanceof Date) {
    const hours = String(value.getHours()).padStart(2, "0");
    const minutes = String(value.getMinutes()).padStart(2, "0");
    const seconds = String(value.getSeconds()).padStart(2, "0");

    return `${hours}:${minutes}:${seconds}`;
  }

  // Excel may return fraction of a day
  if (typeof value === "number") {
    let totalSeconds = Math.round(
      value * 24 * 60 * 60
    );

    totalSeconds =
      totalSeconds % (24 * 60 * 60);

    const hours = Math.floor(
      totalSeconds / 3600
    );

    const minutes = Math.floor(
      (totalSeconds % 3600) / 60
    );

    const seconds =
      totalSeconds % 60;

    return `${String(hours).padStart(2, "0")}:${String(
      minutes
    ).padStart(2, "0")}:${String(seconds).padStart(2, "0")}`;
  }

  // Excel may return a string
  if (typeof value === "string") {
    const text = value.trim();

    // HH:MM or HH:MM:SS
    const match = text.match(
      /(\d{1,2}):(\d{2})(?::(\d{2}))?/
    );

    if (match) {
      const hours = String(
        Number(match[1])
      ).padStart(2, "0");

      const minutes = match[2];

      const seconds =
        match[3] || "00";

      return `${hours}:${minutes}:${seconds}`;
    }

    // Try parsing date-like string
    const parsed = new Date(text);

    if (!Number.isNaN(parsed.getTime())) {
      const hours = String(
        parsed.getHours()
      ).padStart(2, "0");

      const minutes = String(
        parsed.getMinutes()
      ).padStart(2, "0");

      const seconds = String(
        parsed.getSeconds()
      ).padStart(2, "0");

      return `${hours}:${minutes}:${seconds}`;
    }
  }

  return null;
}

// =====================================================
// Normalize Header
// =====================================================

function normalizeHeader(value) {
  return String(value ?? "")
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();
}

// =====================================================
// Get Cell Value
// =====================================================

function getCellValue(cell) {
  if (!cell) return null;

  const value = cell.value;

  if (
    value &&
    typeof value === "object" &&
    value.text
  ) {
    return value.text;
  }

  if (
    value &&
    typeof value === "object" &&
    value.result !== undefined
  ) {
    return value.result;
  }

  return value;
}

// =====================================================
// Column Map
// =====================================================

const COLUMN_MAP = {
  date: [
    "التاريخ",
    "التاريخ ميلادي",
    "date",
  ],

  period: [
    "الفترة",
    "الفتره",
    "period",
  ],

  timeFrom: [
    "من",
    "timefrom",
    "time from",
  ],

  timeTo: [
    "إلى",
    "الى",
    "timeto",
    "time to",
  ],

  day: [
    "اليوم",
    "day",
  ],

  professor: [
    "prof",
    "professor",
    "أستاذ",
    "اسم الأستاذ",
    "اسم الاستاذ",
  ],

  course: [
    "course",
    "course_desc",
    "المادة",
    "اسم المقرر",
    "course name",
    "course (name-link)",
  ],

  crn: [
    "crn",
  ],

  courseId: [
    "course_id",
    "course id",
    "courseid",
    "رقم المقرر",
  ],

  syncLink: [
    "sync link",
    "رابط",
    "link",
  ],

  sessionName: [
    "name of session",
  ],
};

function mapHeader(header) {
  const normalized =
    normalizeHeader(header);

  for (
    const [key, variants]
    of Object.entries(COLUMN_MAP)
  ) {
    for (
      const variant of variants
    ) {
      if (
        normalized ===
        normalizeHeader(variant)
      ) {
        return key;
      }
    }
  }

  return null;
}

// =====================================================
// Read Worksheet
// =====================================================

function readWorksheet(worksheet) {
  const rows = [];

  // Excel file has empty rows 1 and 2.
  // Headers are on row 3.
  const HEADER_ROW_NUMBER = 3;

  const headerRow =
    worksheet.getRow(
      HEADER_ROW_NUMBER
    );

  const headerMap = {};

  console.log(
    `📌 Excel headers (row ${HEADER_ROW_NUMBER}):`,
    headerRow.values
  );

  // ---------------------------------------------------
  // Build Header Map
  // ---------------------------------------------------

  headerRow.eachCell(
    {
      includeEmpty: false,
    },
    (cell, colNumber) => {
      const originalHeader =
        getCellValue(cell);

      const key =
        mapHeader(originalHeader);

      if (key) {
        headerMap[colNumber] =
          key;

        console.log(
          `   Column ${colNumber}: "${originalHeader}" → ${key}`
        );
      }
    }
  );

  console.log(
    "📌 Excel header map:",
    headerMap
  );

  // ---------------------------------------------------
  // Read Data Rows
  // ---------------------------------------------------

  worksheet.eachRow(
    {
      includeEmpty: false,
    },
    (row, rowNumber) => {
      // Skip rows 1, 2 and header row 3
      if (
        rowNumber <=
        HEADER_ROW_NUMBER
      ) {
        return;
      }

      const rowData = {};

      row.eachCell(
        {
          includeEmpty: false,
        },
        (cell, colNumber) => {
          const key =
            headerMap[colNumber];

          if (key) {
            rowData[key] =
              getCellValue(cell);
          }
        }
      );

      if (
        Object.keys(rowData).length > 0
      ) {
        rows.push(rowData);
      }
    }
  );

  console.log(
    `📊 Worksheet data rows read: ${rows.length}`
  );

  return {
    rows,
    headerMap,
  };
}

// =====================================================
// Bulk Insert Helper
// =====================================================

async function bulkInsert(
  client,
  table,
  columns,
  rows,
  chunkSize = 500
) {
  if (!rows.length) {
    return [];
  }

  const insertedRows = [];

  for (
    let start = 0;
    start < rows.length;
    start += chunkSize
  ) {
    const chunk =
      rows.slice(
        start,
        start + chunkSize
      );

    const values = [];
    const placeholders = [];

    let parameterIndex = 1;

    for (
      const row of chunk
    ) {
      const rowPlaceholders = [];

      for (
        const value of row
      ) {
        values.push(value);

        rowPlaceholders.push(
          `$${parameterIndex++}`
        );
      }

      placeholders.push(
        `(${rowPlaceholders.join(",")})`
      );
    }

    const query = `
      INSERT INTO ${table}(
        ${columns.join(",")}
      )
      VALUES
        ${placeholders.join(",")}
      RETURNING *
    `;

    const result =
      await client.query(
        query,
        values
      );

    insertedRows.push(
      ...result.rows
    );
  }

  return insertedRows;
}

// =====================================================
// Process Excel Buffer
// =====================================================

async function processExcel(buffer) {
  const workbook =
    new ExcelJS.Workbook();

  await workbook.xlsx.load(buffer);

  if (
    !workbook.worksheets.length
  ) {
    throw new Error(
      "No worksheets found in Excel file"
    );
  }

  const worksheet =
    workbook.worksheets[0];

  console.log(
    `📄 Reading worksheet: ${worksheet.name}`
  );

  const {
    rows,
    headerMap,
  } =
    readWorksheet(worksheet);

  if (!rows.length) {
    throw new Error(
      "Excel file contains no data rows"
    );
  }

  // ===================================================
  // Validate Rows
  // ===================================================

  const requiredFields = [
    "date",
    "period",
    "timeFrom",
    "timeTo",
    "crn",
    "professor",
  ];

  const invalidRows = [];
  const validRows = [];

  rows.forEach(
    (row, index) => {
      const missingFields =
        requiredFields.filter(
          (field) =>
            row[field] ===
              undefined ||
            row[field] ===
              null ||
            String(row[field])
              .trim() === ""
        );

      if (
        missingFields.length
      ) {
        invalidRows.push({
          row: index + 4,
          missing:
            missingFields,
        });
      } else {
        validRows.push(row);
      }
    }
  );

  console.log(
    `📊 Excel rows: ${rows.length}`
  );

  console.log(
    `✅ Valid rows: ${validRows.length}`
  );

  console.log(
    `⚠️ Invalid rows: ${invalidRows.length}`
  );

  // ===================================================
  // Create Batch
  // ===================================================

  const batchId =
    uuidv4();

  const client =
    await pool.connect();

  try {
    await client.query(
      "BEGIN"
    );

    // =================================================
    // 1. Normalize all rows first
    // =================================================

    const normalizedRows = [];

    for (
      const row of validRows
    ) {
      const date =
        excelDateToJSDate(
          row.date
        );

      if (!date) {
        console.warn(
          "⚠️ Invalid date detected:",
          row.date
        );

        continue;
      }

      const timeFrom =
        normalizeTime(
          row.timeFrom
        );

      const timeTo =
        normalizeTime(
          row.timeTo
        );

      if (
        !timeFrom ||
        !timeTo
      ) {
        console.warn(
          "⚠️ Invalid time detected:",
          {
            timeFrom:
              row.timeFrom,
            timeTo:
              row.timeTo,
          }
        );

        continue;
      }

      const professorName =
        row.professor
          ? String(
              row.professor
            ).trim()
          : null;

      const courseName =
        row.course
          ? String(
              row.course
            ).trim()
          : null;

      const crn =
        row.crn
          ? String(
              row.crn
            ).trim()
          : null;

      const courseIdText =
        row.courseId
          ? String(
              row.courseId
            ).trim()
          : null;

      normalizedRows.push({
        original: row,
        date,
        timeFrom,
        timeTo,
        professorName,
        courseName,
        crn,
        courseIdText,
        day: row.day
          ? String(
              row.day
            ).trim()
          : null,
        period: String(
          row.period
        ).trim(),
        syncLink:
          row.syncLink
            ? String(
                row.syncLink
              ).trim()
            : null,
        sessionName:
          row.sessionName
            ? String(
                row.sessionName
              ).trim()
            : null,
      });
    }

    console.log(
      `📦 Normalized rows: ${normalizedRows.length}`
    );

    // =================================================
    // 2. Load existing professors
    // =================================================

    const professorCache =
      new Map();

    const professorsResult =
      await client.query(`
        SELECT id, name
        FROM professors
      `);

    for (
      const professor
      of professorsResult.rows
    ) {
      professorCache.set(
        String(
          professor.name
        ).trim(),
        professor.id
      );
    }

    console.log(
      `👨‍🏫 Existing professors loaded: ${professorCache.size}`
    );

    // =================================================
    // 3. Create missing professors
    // =================================================

    const uniqueProfessorNames =
      [
        ...new Set(
          normalizedRows
            .map(
              (row) =>
                row.professorName
            )
            .filter(Boolean)
        ),
      ];

    const newProfessorNames =
      uniqueProfessorNames.filter(
        (name) =>
          !professorCache.has(
            name
          )
      );

    console.log(
      `👨‍🏫 New professors: ${newProfessorNames.length}`
    );

    for (
      const professorName
      of newProfessorNames
    ) {
      const result =
        await client.query(
          `
          INSERT INTO professors(name)
          VALUES($1)
          ON CONFLICT(name)
          DO UPDATE
          SET name = EXCLUDED.name
          RETURNING id
          `,
          [professorName]
        );

      professorCache.set(
        professorName,
        result.rows[0].id
      );
    }

    // =================================================
    // 4. Load existing courses
    // =================================================

    const courseCache =
      new Map();

    const coursesResult =
      await client.query(`
        SELECT
          id,
          name,
          crn,
          course_id
        FROM courses
      `);

    for (
      const course
      of coursesResult.rows
    ) {
      if (
        course.name &&
        !courseCache.has(
          String(
            course.name
          ).trim()
        )
      ) {
        courseCache.set(
          String(
            course.name
          ).trim(),
          course.id
        );
      }
    }

    console.log(
      `📚 Existing courses loaded: ${courseCache.size}`
    );

    // =================================================
    // 5. Create missing courses
    // =================================================

    const uniqueCourses =
      new Map();

    for (
      const row
      of normalizedRows
    ) {
      if (
        row.courseName &&
        !uniqueCourses.has(
          row.courseName
        )
      ) {
        uniqueCourses.set(
          row.courseName,
          {
            crn: row.crn,
            courseIdText:
              row.courseIdText,
          }
        );
      }
    }

    let newCoursesCount = 0;

    for (
      const [
        courseName,
        courseData,
      ]
      of uniqueCourses.entries()
    ) {
      if (
        courseCache.has(
          courseName
        )
      ) {
        continue;
      }

      const result =
        await client.query(
          `
          INSERT INTO courses(
            name,
            crn,
            course_id
          )
          VALUES($1,$2,$3)
          RETURNING id
          `,
          [
            courseName,
            courseData.crn,
            courseData.courseIdText,
          ]
        );

      courseCache.set(
        courseName,
        result.rows[0].id
      );

      newCoursesCount++;
    }

    console.log(
      `📚 New courses: ${newCoursesCount}`
    );

    // =================================================
    // 6. Bulk Insert Raw Sessions
    // =================================================

    const rawSessionRows =
      normalizedRows.map(
        (row) => [
          batchId,
          row.day,
          row.date,
          row.period,
          row.timeFrom,
          row.timeTo,
          row.syncLink,
          row.professorName,
          row.courseName,
          row.crn,
          row.courseIdText,
          row.sessionName,
        ]
      );

    console.log(
      `💾 Inserting ${rawSessionRows.length} raw sessions...`
    );

    const rawSessions =
      await bulkInsert(
        client,
        "raw_sessions",
        [
          "excel_batch_id",
          "day_name",
          "date",
          "period_label",
          "time_from",
          "time_to",
          "sync_link",
          "professor_name",
          "course_text",
          "crn",
          "course_id",
          "session_name",
        ],
        rawSessionRows,
        500
      );

    console.log(
      `✅ Raw sessions inserted: ${rawSessions.length}`
    );

    // =================================================
    // 7. Bulk Insert Session Links
    // =================================================

    const sessionLinkRows =
      [];

    for (
      let i = 0;
      i < rawSessions.length;
      i++
    ) {
      const rawSession =
        rawSessions[i];

      const sourceRow =
        normalizedRows[i];

      const professorId =
        sourceRow.professorName
          ? professorCache.get(
              sourceRow.professorName
            )
          : null;

      const courseDbId =
        sourceRow.courseName
          ? courseCache.get(
              sourceRow.courseName
            )
          : null;

      if (
        professorId ||
        courseDbId
      ) {
        sessionLinkRows.push([
          rawSession.id,
          professorId,
          courseDbId,
          sourceRow.crn,
        ]);
      }
    }

    console.log(
      `🔗 Inserting ${sessionLinkRows.length} session links...`
    );

    await bulkInsert(
      client,
      "session_links",
      [
        "raw_session_id",
        "professor_id",
        "course_id",
        "crn",
      ],
      sessionLinkRows,
      500
    );

    console.log(
      `✅ Session links inserted`
    );

    // =================================================
    // 8. Create Session Groups
    // =================================================

    const groupsResult =
      await client.query(
        `
        SELECT
          rs.excel_batch_id,
          rs.date,
          rs.period_label,
          rs.time_from,
          rs.time_to,
          rs.crn,
          sl.professor_id,
          COUNT(*)::int AS sessions

        FROM raw_sessions rs

        LEFT JOIN session_links sl
          ON sl.raw_session_id =
             rs.id

        WHERE
          rs.excel_batch_id = $1

        GROUP BY
          rs.excel_batch_id,
          rs.date,
          rs.period_label,
          rs.time_from,
          rs.time_to,
          rs.crn,
          sl.professor_id

        ORDER BY
          rs.date,
          rs.time_from,
          rs.time_to
        `,
        [batchId]
      );

    console.log(
      `📦 Creating ${groupsResult.rows.length} session groups`
    );

    // =================================================
    // 9. Bulk Insert Session Groups
    // =================================================

    const sessionGroupRows =
      groupsResult.rows.map(
        (group) => [
          group.excel_batch_id,
          group.date,
          group.period_label,
          group.time_from,
          group.time_to,
          group.crn,
          group.professor_id,
          1,
          group.sessions,
        ]
      );

    if (
      sessionGroupRows.length
    ) {
      await bulkInsert(
        client,
        "session_groups",
        [
          "excel_batch_id",
          "date",
          "period_label",
          "time_from",
          "time_to",
          "crn",
          "professor_id",
          "required_supervisors",
          "sessions",
        ],
        sessionGroupRows,
        500
      );
    }

    console.log(
      `✅ Session groups inserted: ${sessionGroupRows.length}`
    );

    // =================================================
    // 10. Commit
    // =================================================

    await client.query(
      "COMMIT"
    );

    console.log(
      "✅ Database transaction committed"
    );

    // =================================================
    // Add invalid flags
    // =================================================

    const invalidMap =
      new Map();

    for (
      const invalid
      of invalidRows
    ) {
      invalidMap.set(
        invalid.row,
        invalid
      );
    }

    const allRows =
      rows.map(
        (row, index) => {
          const excelRowNumber =
            index + 4;

          const invalid =
            invalidMap.get(
              excelRowNumber
            );

          if (invalid) {
            return {
              ...row,
              __invalid: true,
              missing:
                invalid.missing,
            };
          }

          return row;
        }
      );

    return {
      batchId,

      rows: allRows,

      validRows:
        validRows.length,

      invalidRows:
        invalidRows.length,

      sessionGroups:
        groupsResult.rows.length,

      headerMap,
    };

  } catch (error) {
    console.error(
      "❌ processExcel failed:",
      error
    );

    try {
      await client.query(
        "ROLLBACK"
      );
    } catch (
      rollbackError
    ) {
      console.error(
        "❌ Rollback failed:",
        rollbackError
      );
    }

    throw error;

  } finally {
    client.release();
  }
}

// =====================================================
// Upload Excel
// =====================================================

async function uploadExcel(
  req,
  res
) {
  try {
    if (!req.file) {
      return res.status(400).json({
        success: false,
        message:
          "No file uploaded",
      });
    }

    console.log(
      `📥 Uploading Excel: ${req.file.originalname}`
    );

    const result =
      await processExcel(
        req.file.buffer
      );

    console.log(
      "✅ Excel uploaded successfully"
    );

    console.log(
      `🆔 Batch ID: ${result.batchId}`
    );

    console.log(
      `📦 Session groups: ${result.sessionGroups}`
    );

    return res.json({
      success: true,

      message:
        "Excel file uploaded and processed successfully",

      data: result.rows,

      inserted:
        result.validRows,

      skipped:
        result.invalidRows,

      session_groups:
        result.sessionGroups,

      excel_batch_id:
        result.batchId,
    });

  } catch (err) {
    console.error(
      "❌ Excel upload error:",
      err.stack || err
    );

    return res.status(500).json({
      success: false,

      message:
        err.message ||
        "Failed to process Excel file",
    });
  }
}

// =====================================================
// Upload Excel From URL
// =====================================================

async function uploadExcelFromUrl(
  req,
  res
) {
  try {
    const { url } =
      req.body;

    if (!url) {
      return res.status(400).json({
        success: false,
        message:
          "URL is required",
      });
    }

    console.log(
      `🌐 Fetching Excel from URL: ${url}`
    );

    const response =
      await fetch(url);

    if (!response.ok) {
      return res.status(400).json({
        success: false,
        message:
          "Failed to fetch Excel file",
      });
    }

    const buffer =
      Buffer.from(
        await response.arrayBuffer()
      );

    const result =
      await processExcel(
        buffer
      );

    console.log(
      "✅ Excel imported successfully"
    );

    return res.json({
      success: true,

      message:
        "Excel file imported successfully",

      data: result.rows,

      inserted:
        result.validRows,

      skipped:
        result.invalidRows,

      session_groups:
        result.sessionGroups,

      excel_batch_id:
        result.batchId,
    });

  } catch (err) {
    console.error(
      "❌ Excel URL upload error:",
      err.stack || err
    );

    return res.status(500).json({
      success: false,

      message:
        err.message ||
        "Failed to process Excel file",
    });
  }
}

// =====================================================
// Generate Excel
// =====================================================

async function generateExcel(
  req,
  res
) {
  return res.json({
    success: true,

    message:
      "generateExcel endpoint is available",
  });
}

// =====================================================
// Get Raw Sessions
// =====================================================

async function getRawSessions(
  req,
  res
) {
  try {
    const { batchId } =
      req.params;

    if (!batchId) {
      return res.status(400).json({
        success: false,
        message:
          "batchId is required",
      });
    }

    const result =
      await pool.query(
        `
        SELECT
          rs.*,
          p.name AS professor,
          c.name AS course_name

        FROM raw_sessions rs

        LEFT JOIN session_links sl
          ON sl.raw_session_id =
             rs.id

        LEFT JOIN professors p
          ON p.id = sl.professor_id

        LEFT JOIN courses c
          ON c.id = sl.course_id

        WHERE
          rs.excel_batch_id = $1

        ORDER BY
          rs.date,
          rs.period_label,
          rs.time_from
        `,
        [batchId]
      );

    return res.json({
      success: true,

      data:
        result.rows,
    });

  } catch (err) {
    return handleError(
      res,
      err,
      "Error fetching raw sessions"
    );
  }
}

// =====================================================
// Error Helper
// =====================================================

function handleError(
  res,
  err,
  message
) {
  console.error(
    message,
    err
  );

  return res.status(500).json({
    success: false,

    message:
      err.message ||
      message,
  });
}

// =====================================================
// Exports
// =====================================================

module.exports = {
  uploadExcel,
  uploadExcelFromUrl,
  generateExcel,
  getRawSessions,
};