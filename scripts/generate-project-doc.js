const fs = require("fs");
const path = require("path");
const {
  AlignmentType,
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType
} = require("docx");

const OUTPUT_FILE = path.join(
  __dirname,
  "..",
  "documents",
  "Customs_Billing_Portal_Documentation.docx"
);

function heading(text, level) {
  return new Paragraph({
    text,
    heading: level,
    spacing: { after: 160 }
  });
}

function body(text) {
  return new Paragraph({
    children: [new TextRun({ text })],
    spacing: { after: 120 }
  });
}

function bullet(text) {
  return new Paragraph({
    text,
    bullet: { level: 0 },
    spacing: { after: 80 }
  });
}

function spacer() {
  return new Paragraph({ text: "", spacing: { after: 120 } });
}

function createTimeImpactTable() {
  const headers = [
    "Task / Step",
    "Before App (mins)",
    "After App (mins)",
    "Time Saved (mins)",
    "Remarks"
  ];

  const headerRow = new TableRow({
    children: headers.map(
      (title) =>
        new TableCell({
          children: [
            new Paragraph({
              children: [new TextRun({ text: title, bold: true })],
              alignment: AlignmentType.CENTER
            })
          ]
        })
    )
  });

  const blankRows = [];
  for (let i = 0; i < 5; i++) {
    blankRows.push(
      new TableRow({
        children: [
          new TableCell({ children: [new Paragraph("")] }),
          new TableCell({ children: [new Paragraph("")] }),
          new TableCell({ children: [new Paragraph("")] }),
          new TableCell({ children: [new Paragraph("")] }),
          new TableCell({ children: [new Paragraph("")] })
        ]
      })
    );
  }

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [headerRow, ...blankRows]
  });
}

function addFeatureSection(children, feature) {
  children.push(heading(feature.name, HeadingLevel.HEADING_2));
  children.push(body("Feature Objective"));
  feature.objective.forEach((line) => children.push(bullet(line)));
  children.push(body("Current Behavior"));
  feature.currentBehavior.forEach((line) => children.push(bullet(line)));
  children.push(body("Expected Results"));
  feature.expectedResults.forEach((line) => children.push(bullet(line)));
  children.push(body("Time Impact Table (to be filled by operations)"));
  children.push(createTimeImpactTable());
  children.push(spacer());
}

async function generateDocument() {
  const features = [
    {
      name: "1. Excel Merger",
      objective: [
        "Combine multiple Excel files quickly into one consistent output.",
        "Reduce repetitive manual copy and paste tasks for customs clerks."
      ],
      currentBehavior: [
        "Accepts multiple .xlsx files and merges records into a single downloadable merged.xlsx.",
        "Supports row slicing defaults (first file: 4 rows, other files: 5 rows) and optional source filename column.",
        "Generates an on-screen summary including total rows, total duty, total GST, and brokerage value counts."
      ],
      expectedResults: [
        "Faster file consolidation with fewer manual handling steps.",
        "More consistent merged outputs across repeated daily runs.",
        "Improved visibility of totals and brokerage value distributions for quick checks."
      ]
    },
    {
      name: "2. D/T Header File Modifier",
      objective: [
        "Automate repetitive update steps for Duties Header files using SFTP source data.",
        "Prevent manual insertion mistakes when adding new CCN rows."
      ],
      currentBehavior: [
        "Takes one SFTP source file and one target Duties Header file, then inserts only new CCNs.",
        "Maps source data into required target columns and auto-populates fixed values (for example CLVS and DDP).",
        "Normalizes numeric ranges for Value for Duty through Exchange Rate and auto-downloads an updated output file."
      ],
      expectedResults: [
        "Reduced manual data entry effort and lower risk of missed or duplicate CCNs.",
        "Consistent row structure and numeric formatting in every processed file.",
        "Faster completion of repetitive header maintenance tasks."
      ]
    },
    {
      name: "3. Header/Item Analyzer",
      objective: [
        "Provide quick validation checks on Duties Header and optional Duties Item files.",
        "Help clerks identify exceptions and mismatches earlier in the workflow."
      ],
      currentBehavior: [
        "Analyzes key counts and exception lists, including empty brokerage, empty value for duty, and low-value duty or GST anomalies.",
        "Computes total duty and GST for header and optional item file.",
        "Displays totals comparison (matched or not matched) when both files are provided."
      ],
      expectedResults: [
        "Faster review and anomaly detection before downstream handoff.",
        "Better accuracy through early mismatch detection between header and item totals.",
        "Reduced time spent on manual reconciliation checks."
      ]
    },
    {
      name: "4. Candata to Gets Format",
      objective: [
        "Convert Candata Duties Header and Candata Duties Item files into GETS-formatted outputs.",
        "Standardize conversion steps that are repetitive and prone to manual formatting errors."
      ],
      currentBehavior: [
        "Requires both Candata Duties Header and Candata Duties Item files and produces two GETS output files.",
        "Builds normalized output structures for header and item data, including report metadata and sorted rows.",
        "Current output rules include blank client name (A2) and blank brokerage totals in header output rows."
      ],
      expectedResults: [
        "Reliable and repeatable conversion to GETS format with minimal manual intervention.",
        "Reduced rework caused by inconsistent field mapping and formatting.",
        "Faster turnaround for producing ready-to-use header and item files."
      ]
    }
  ];

  const children = [];
  children.push(
    new Paragraph({
      text: "Customs Billing Portal (Excel Merger) - Project Documentation",
      heading: HeadingLevel.TITLE,
      alignment: AlignmentType.CENTER,
      spacing: { after: 240 }
    })
  );
  children.push(
    body(
      "This document summarizes the purpose, key features, and expected operational results of the Excel Merger project (Customs Billing Portal). The platform is intended to support customs clerks by automating repetitive, time-consuming Excel tasks and improving consistency in daily processing."
    )
  );

  children.push(heading("Project Purpose", HeadingLevel.HEADING_1));
  children.push(
    bullet(
      "Provide convenience to clerks by automating repetitive spreadsheet tasks used in customs workflows."
    )
  );
  children.push(
    bullet(
      "Reduce manual processing time, repetitive clicking, and copy or paste activity across common file operations."
    )
  );
  children.push(
    bullet(
      "Improve consistency and quality of outputs by enforcing repeatable transformation and validation steps."
    )
  );
  children.push(
    bullet(
      "Support faster handling of daily operational volume while reducing avoidable data handling errors."
    )
  );
  children.push(spacer());

  children.push(heading("Features and Expected Results", HeadingLevel.HEADING_1));
  children.push(
    body(
      "Each feature below includes a blank time-impact table. These fields are intentionally left empty for operations to fill with real before and after timing measurements."
    )
  );

  features.forEach((feature) => addFeatureSection(children, feature));

  children.push(heading("Conclusion", HeadingLevel.HEADING_1));
  children.push(
    body(
      "The Customs Billing Portal is designed to improve day-to-day efficiency for customs clerks through automation of repetitive Excel tasks. Expected outcomes include faster processing, more consistent outputs, and reduced manual error risk across merging, modification, analysis, and conversion workflows."
    )
  );
  children.push(
    body(
      "All time-impact fields in this document are intentionally blank and should be completed by operations personnel using observed process timings."
    )
  );

  const document = new Document({
    title: "Customs Billing Portal (Excel Merger) - Project Documentation",
    description: "Purpose, features, expected results, and time-impact tables for Customs Billing Portal.",
    creator: "Customs Billing Portal Team",
    lastModifiedBy: "Customs Billing Portal Team",
    sections: [
      {
        properties: {},
        children
      }
    ]
  });

  fs.mkdirSync(path.dirname(OUTPUT_FILE), { recursive: true });
  const buffer = await Packer.toBuffer(document);
  fs.writeFileSync(OUTPUT_FILE, buffer);
  console.log(`Generated: ${OUTPUT_FILE}`);
}

generateDocument().catch((error) => {
  console.error("Failed to generate project documentation:", error);
  process.exitCode = 1;
});
