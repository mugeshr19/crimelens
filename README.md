# 🦅 Crime Lens

> **An AI-first Crime Intelligence Operating System for the Karnataka State Police**  
> Real-time hotspots · criminal-network graphs · predictive patrols · tamper-evident governance — on the web and in an officer’s pocket.

---

## 📁 Project Structure

This repository contains the complete CrimeLens platform:

- **[`CrimeLens_Full_Project/`](./CrimeLens_Full_Project)** — Main full-stack application
  - **`app/`** — React Native / Expo mobile application
  - **`web/`** — React / Vite web dashboard & live map
  - **`server/`** — Node.js & TypeScript backend API
  - **`database/`** — Drizzle ORM schema and migrations
  - **`components/`** — Shared UI component library
  - **`start.ps1`** — PowerShell script to launch both backend and frontend

For detailed architecture, setup instructions, and feature blueprints, refer to the documentation inside **[`CrimeLens_Full_Project/README.md`](./CrimeLens_Full_Project/README.md)**.

---

### Quick Start

To launch the project locally:

```powershell
cd CrimeLens_Full_Project
.\start.ps1
```

- **Frontend:** http://localhost:3000
- **Backend:** http://localhost:5000
- **Dashboard:** http://localhost:3000/dashboard
- **Live Map:** http://localhost:3000/map
