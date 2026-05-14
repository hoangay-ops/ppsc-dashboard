# AI Operating System for Dashboard Development

## Purpose
This document defines how AI assistants should behave when working on this dashboard codebase.

The goal is to ensure:

minimal breaking changes
consistent architecture
safe feature additions
reuse of existing code patterns
prevention of repeated past failures


---

## 1. System Definition

This is an existing production dashboard built with:

Python (Flask backend)
HTML templates (frontend)
Bootstrap styling
Session-based filtering system
Chart rendering via shared helper functions


This system already works and must NOT be rewritten unless explicitly requested.

---

## 2. Core Principles (NON-NEGOTIABLE)

When modifying this system:

### Preserve Existing Functionality

Do not break working features
Do not remove working routes or logic
Do not replace working architecture


### Minimal Change Policy

Only change what is necessary for the feature
Avoid refactoring unrelated code
Avoid “clean up” unless requested


### Reuse Existing Code

Always reuse helper functions when available
Do not duplicate logic
Do not introduce new patterns if existing ones work


---

## 3. Required Workflow Before Any Code Change

Before writing any code, the AI MUST:


Identify relevant files
Explain how the current system works
Identify dependencies and impacted components
List risks or potential breakpoints
Propose a minimal implementation plan


NO code should be written before completing these steps.

---

## 4. High-Risk Areas (DO NOT BREAK)

These areas are fragile and must be handled carefully:


Filtering system (session-based state)
Chart rendering logic (must use existing helpers)
Flask routes (must not be renamed)
Shared UI components (sidebar, navigation)
Data transformation logic before visualization


---

## 5. Known Failure Patterns (DO NOT REPEAT)

### Filter System Breakage
Cause:

Duplicate filter IDs or recreated state objects


Rule:

Always reuse existing filter state structure


---

### Chart Breakage
Cause:

Bypassing chart helper functions


Rule:

ALL charts must use existing chart helper utilities


---

### Route Breakage
Cause:

Renaming or restructuring Flask routes


Rule:

Never rename routes unless explicitly required


---

## 6. Change Safety Rules

Before finalizing any code:


Confirm no unrelated files were modified
Confirm no duplication of logic was introduced
Confirm existing UI structure is preserved
Confirm no breaking changes to filtering or charts
Confirm imports and dependencies are valid


---

## 7. Output Requirements

When proposing changes:


Prefer explanation BEFORE code
Provide file-by-file impact list
Provide step-by-step implementation plan
Keep changes small and reversible


---

## 8. Definition of “Success”

A successful change is one where:

existing dashboard still works unchanged
only intended feature was added
no architectural rewrites occurred
no regressions in filters or charts