# Automated SLA Monitoring & Staffing Optimization System

## Overview

Built an automated SLA monitoring and workforce optimization system integrating ERPNext service operations data to evaluate field performance, identify staffing gaps, and automate reporting workflows.

The system processed multiple field service job categories, computed SLA adherence metrics, and generated region-level staffing recommendations using automated analytics pipelines.

---

## Business Problem

Field service SLA tracking and workforce planning were previously dependent on manual Excel-based analysis across multiple service regions.

Key operational challenges included:

- No centralized SLA monitoring process
- Manual turnaround-time calculations
- Delayed identification of SLA breaches
- Difficulty estimating staffing shortages
- No historical trend comparison
- High reporting effort for regional managers

The reporting cycle required significant manual consolidation and operational follow-up.

---

## Solution

Developed a Python-based automation system that:

- Integrated ERPNext service job data through APIs
- Processed multiple field service job categories
- Computed Door-to-Close turnaround time
- Classified SLA performance into:
  - On Time
  - Near SLA
  - Late
- Built staffing optimization logic at sub-region level
- Generated automated Excel reports
- Maintained JSON-based historical trend records
- Distributed reports automatically via SMTP email workflows

---

## SLA Classification Logic

Service jobs were evaluated using predefined SLA thresholds based on job category.

Classification Categories:

- On Time
- Near SLA
- Late

The logic enabled operational visibility into SLA risk and service bottlenecks.

---

## Staffing Optimization Logic

Required staffing estimates were calculated using SLA achievement percentages.

Staffing Formula:

Required Staff = Current Staff × 0.95 ÷ SLA %

The model identified workforce gaps and operational load imbalance at sub-region level.

---

## System Workflow

```text
ERPNext APIs
      ↓
Service Job Extraction
      ↓
Door-to-Close Time Calculation
      ↓
SLA Classification Engine
      ↓
Regional SLA Aggregation
      ↓
Staffing Optimization Logic
      ↓
Historical Trend Storage (JSON)
      ↓
Excel Report Generation
      ↓
SMTP-Based Automated Distribution
```
## Key Features

- ERPNext REST API integration
- Automated SLA computation
- Door-to-Close turnaround analytics
- Job-type-specific SLA thresholds
- Workforce gap analysis
- Historical JSON trend database
- Automated Excel reporting
- SMTP-based report distribution
- Regional and sub-region performance tracking

---

## Engineering Challenges Solved

- Standardized SLA calculations across multiple job categories
- Handled varying turnaround-time thresholds dynamically
- Built reusable staffing optimization logic
- Automated historical trend persistence using JSON
- Reduced manual dependency for regional reporting workflows
- Scaled reporting pipelines for multiple operational regions

---

## Tech Stack

- Python
- Pandas
- Requests
- JSON
- OpenPyXL
- SMTP
- Windows Task Scheduler

---

## Project Metrics

| Metric | Value |
|---|---|
| Job Categories Processed | 4 |
| Reporting Scope | Multi-Region |
| Historical Tracking | JSON Database |
| Report Delivery | Automated SMTP Distribution |
| Staffing Analysis Level | Sub-Region |
| Manual Effort Reduction | 8–10 hrs per monthly cycle |

---

## Scalability Considerations

The system architecture was designed to support:

- Additional service regions
- New SLA job categories
- Dynamic staffing calculations
- Automated recurring report execution
- Historical trend expansion
- Scalable reporting workflows

---

## Future Enhancements

- Interactive SLA dashboards
- Real-time SLA alerting
- Predictive staffing models
- SQL-based historical storage
- Power BI integration
- Cloud-hosted reporting pipelines

---

## Business Impact

- Improved SLA visibility across field operations
- Enabled proactive staffing decisions
- Reduced operational reporting effort
- Standardized SLA evaluation methodology
- Automated recurring reporting and distribution workflows

---

## Disclaimer

This repository is a portfolio representation of work completed during an internship.

Source code, credentials, internal ERP configurations, and business-sensitive operational logic have been excluded to comply with company confidentiality policies.
