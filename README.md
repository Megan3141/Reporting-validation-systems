# Dashboard Validation System

A Python validation tool for checking recurring Excel-based reporting dashboards before publication.

## Overview

Recurring reporting dashboards often rely on manually maintained Excel files, copied data, linked formulas, and fixed row structures. Small layout changes, missing rows, or blank values can lead to incorrect outputs.

This project provides a reusable validation layer that checks dashboard inputs and outputs before reports are finalised.

## What It Checks

- Missing expected rows
- Unexpected new rows
- Zero or blank values in key metrics
- Changes in dashboard structure between reporting periods
- Differences between current week and prior week layouts
- Potential mapping issues before dashboard publication

## Why This Matters

The tool reduces the risk of publishing incorrect reporting outputs by flagging issues before dashboards are saved, exported, or shared.

## Example Workflow

```text
Current Dashboard Data
        ↓
Prior Period Comparison
        ↓
Validation Checks
        ↓
Exception Report
        ↓
Manual Review / Fix
