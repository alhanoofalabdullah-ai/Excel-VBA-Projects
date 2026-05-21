# Financial Audit & Compliance Tracker – Excel VBA

A professional Excel VBA solution designed to support financial audit activities, compliance tracking, finding management, risk scoring, corrective action follow-up, and executive reporting.

This project simulates a real-world internal audit and financial controls tool used to record audit findings, assess risk, assign corrective actions, monitor due dates, and generate management-ready summaries.

---

## Overview

The **Financial Audit & Compliance Tracker** is a macro-enabled Excel workbook built for finance, audit, compliance, and governance workflows.

It enables users to:
- Record financial audit findings in a structured register
- Auto-generate unique Audit IDs
- Assign owners and due dates for corrective actions
- Classify observations by audit area, severity, and financial impact
- Track open, closed, overdue, and high-risk issues
- Refresh dashboards automatically
- Export executive summaries to PDF
- Maintain an activity log for traceability

---

## Key Features

### 1. Structured Finding Register
Capture findings with standardized fields such as:
- Audit ID
- Audit Area
- Department
- Finding Title
- Detailed Observation
- Risk Rating
- Financial Impact
- Recommendation
- Action Owner
- Due Date
- Status

### 2. Automated Controls via VBA
The VBA engine handles:
- automatic ID generation
- data validation
- dashboard refresh
- overdue flagging
- PDF report export
- activity logging
- quick record insertion

### 3. Executive Dashboard
The dashboard provides high-level visibility into:
- total findings
- open findings
- closed findings
- overdue findings
- critical findings
- findings by department
- findings by audit area
- findings by status

### 4. Compliance and Follow-Up Tracking
The tool is suitable for:
- internal audit follow-up
- compliance reviews
- finance control testing
- exception monitoring
- remediation tracking

---

## Workbook Structure

| Sheet Name | Purpose |
|-----------|---------|
| Home | Main landing page with navigation buttons |
| Findings_Register | Master table of all audit findings |
| Action_Tracker | Follow-up actions and status tracking |
| Dashboard | KPI cards, charts, and audit summaries |
| Reference_Lists | Drop-down list values and configuration data |
| Executive_Report | Printable management summary |
| Activity_Log | Log of key macro-driven actions |
| Settings | Control values, prefixes, and project options |
| Instructions | User guide and operating notes |

---

## Example Use Cases

- Internal financial audit tracking
- Departmental compliance review logs
- Risk issue monitoring for finance teams
- Remediation tracking for audit recommendations
- Audit committee monthly reporting support

---

## Skills Demonstrated

- Excel VBA
- Macros and automation
- Financial audit workflow design
- Compliance tracking
- Risk classification
- Executive dashboard automation
- PDF reporting
- Structured business tooling

---

## Sample VBA Functions Included

- `GenerateAuditID()`
- `AddNewFinding()`
- `ValidateFindingRow()`
- `RefreshDashboard()`
- `FlagOverdueFindings()`
- `ExportExecutiveReportToPDF()`
- `WriteToActivityLog()`

---

## Suggested Enhancements

Future versions may include:
- UserForm-based entry screen
- Login / role-based access
- Email notification reminders
- Advanced search/filter panel
- Department heatmap
- Auto-archiving for closed findings

---

## Author
Alhanoof Alabdullah
