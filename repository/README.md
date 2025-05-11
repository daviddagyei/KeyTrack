# KeyTrack Web Application

A modern, secure, and user-friendly web application for managing room key distribution, collection, loss, and borrowing in residential or institutional settings. 

**Live Demo:** [https://keytrack-1.onrender.com/](https://keytrack-1.onrender.com/)

---

## Table of Contents
- [Features](#features)
- [Architecture Overview](#architecture-overview)
- [Component Breakdown](#component-breakdown)
  - [Backend](#backend)
  - [Repository Layer](#repository-layer)
  - [Models](#models)
  - [Services](#services)
  - [Web Application (Flask)](#web-application-flask)
  - [Templates & Static Files](#templates--static-files)
  - [Utilities](#utilities)
  - [Excel Demo Data](#excel-demo-data)
- [Authentication](#authentication)
- [Demo Mode](#demo-mode)
- [Deployment](#deployment)
- [Development & Testing](#development--testing)
- [License](#license)

---

## Features
- Caltech-branded authentication (Supabase, Flask-Login, Caltech email enforcement)
- Demo mode for instant exploration (no login required)
- Key actions: collect, return, report lost, borrow spare, return borrowed
- Realistic sample data for demo users
- Excel-based data storage for easy export/import
- Modern, responsive UI (Bootstrap, custom JS)
- Audit log and statistics dashboard

---

## Architecture Overview

```
[User] <-> [Flask Webapp] <-> [Service Layer] <-> [Repository Layer] <-> [Excel File]
                                      |                        |
                                      |                        +-- [Supabase Auth]
                                      +-- [Templates/Static]
```

---

## Component Breakdown

### Backend
- **`webapp/app.py`**: Main Flask app. Handles routing, authentication, session management, and API endpoints for key actions, statistics, and room overviews.
- **`backend/server.py`**: (If present) May provide additional API endpoints or background services.

### Repository Layer
- **`repository/excel_repository.py`**: Reads/writes all room and key action data to `key_distribution.xlsx`. Handles parsing, serialization, and error handling for Excel I/O.
- **`repository/file_manager.py`**: Manages file uploads, downloads, and file metadata.

### Models
- **`models/entities.py`**: Core domain models: `Room`, `KeyAction`, `Student`. Encapsulate business logic for key actions and room state.
- **`models/base.py`**: Base classes and interfaces for repositories and entities.
- **`models/exceptions.py`**: Custom exception types for robust error handling.

### Services
- **`services/key_management_service.py`**: Implements business logic for collecting, returning, borrowing, and reporting lost keys. Ensures all rules are enforced and updates the repository.

### Web Application (Flask)
- **`webapp/app.py`**: Flask routes for login, signup, dashboard, room details, actions, statistics, and demo mode.
- **Session Management**: Uses Flask-Login for user sessions and integrates with Supabase for secure authentication.

### Templates & Static Files
- **`webapp/templates/`**: Jinja2 HTML templates for all pages (login, signup, dashboard, room detail, actions, etc.), with Caltech branding and responsive design.
- **`webapp/static/`**: CSS, JS, and image assets. Includes dashboard charts, room detail interactivity, and form validation.

### Utilities
- **`utils/auth.py`**: Supabase authentication logic, Caltech email enforcement, and session helpers.
- **`utils/logger.py`**: Logging setup for audit trails and debugging.
- **`utils/excel_creator.py`**: Utility for initializing or resetting Excel files.

### Excel Demo Data
- **`key_distribution.xlsx`**: Main data file. Stores all room and key action data. Populated with realistic demo data for demo users.
- **`populate_demo_excel.py`**: Script to generate realistic sample data for demo mode.

---

## Authentication
- **Supabase**: Handles secure user authentication and session management.
- **CaltechEmail Enforcement**: Only users with `@Caltech.edu` emails can register/login (except in demo mode).
- **Flask-Login**: Manages user sessions and access control.

---

## Demo Mode
- No login required. Instantly loads a sample Excel file with realistic key logs and actions.
- Demo users cannot modify real data or access private user features.
- Always resets to a default state for consistent exploration.

---

## Deployment
- **Render.com**: App is hosted at [https://keytrack-1.onrender.com/](https://keytrack-1.onrender.com/)
- **render.yaml**: Deployment configuration for Render (build/start commands, environment variables).
- **.env**: Store secrets and Supabase credentials (never commit real secrets to version control).

---

## Development & Testing
- **Requirements**: See `requirements.txt` for all dependencies (Flask, Supabase, pandas, openpyxl, Faker, etc.)
- **Testing**: Pytest-based tests in `tests/` for all major components (entities, repository, services, exceptions, logger).
- **Scripts**: Use `populate_demo_excel.py` to reset demo data, and `update_excel.py` for data migration/cleanup.
- **Logs**: All actions and errors are logged to `logs/` for auditing and debugging.

---

## Contact
For questions or contributions, please open an issue or pull request on the [GitHub repository](https://github.com/daviddagyei/KeyTrack).




