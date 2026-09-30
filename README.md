# 🌿 Greenscape — AI Landscape Planning & Estimation Assistant

<p align="center">
  <strong>AI-powered landscape planning, plant recommendation, BOQ generation, cost estimation, and site visualization.</strong>
</p>

<p align="center">
  <a href="https://greenscape-landscape-planner-kwhp4zqojnhweior5ubnux.streamlit.app/">
    🚀 <strong>Live Demo</strong>
  </a>
</p>

---

## 🌱 Overview

**Greenscape** is an AI-powered landscape planning and estimation application designed to simplify the process of converting site requirements into a structured landscape plan.

The application combines **Generative AI, plant data, image analysis, estimation logic, and automated Excel generation** into a single web-based workflow.

Instead of manually preparing plant recommendations, quantities, landscape scope, BOQs, and preliminary estimates, Greenscape provides an interactive interface for generating these outputs from project requirements and site images.

### 🎯 Project Goal

The goal of Greenscape is to create a practical AI assistant for:

- Landscape designers
- Landscape contractors
- Garden and nursery businesses
- Architects
- Property developers
- Homeowners
- Landscape planning teams

---

## 🚀 Live Application

### 🔗 Try Greenscape

**[Open the Live Application →](https://greenscape-landscape-planner-kwhp4zqojnhweior5ubnux.streamlit.app/)**

> The application is deployed using Streamlit Community Cloud.

---

# ✨ Key Features

## 🧠 AI Landscape Planning

Enter project information such as:

- Site area
- Location
- Soil conditions
- Sunlight conditions
- Landscape requirements
- Maintenance preferences
- Selected landscape features

Greenscape uses AI to generate a structured landscape planning report.

---

## 🌿 Plant Recommendation System

The application uses a structured plant database to recommend suitable plants based on the project requirements.

The plant database contains information such as:

- Plant name
- Plant category
- Pricing
- Landscape suitability

The recommendations can be used to create a practical **Plant BOQ**.

---

## 📋 Plant BOQ Generation

Greenscape generates a structured Bill of Quantities containing:

| Information | Description |
|---|---|
| Plant | Recommended plant |
| Category | Plant classification |
| Quantity | Required quantity |
| Rate | Unit price |
| Amount | Estimated plant cost |

This helps convert AI recommendations into a more practical estimation format.

---

## 💰 Cost Estimation

Greenscape calculates an estimated project budget based on the selected scope and generated quantities.

The estimation can include applicable landscape components such as:

- Plant material
- Lawn
- Trees
- Hedges
- Flowering plants
- Ornamental plants
- Selected hardscape features
- Irrigation
- Labour
- Other selected project components

The application is designed so that **only selected project scope is considered for the quotation**.

---

## 🏗️ Landscape & Hardscape Scope

Users can select the requirements that apply to their project.

### Landscape Options

Examples include:

- Lawn
- Trees
- Boundary / Hedge
- Flowering Plants
- Ornamental Plants
- Palm / Tropical Theme
- Vastu Plants
- Soil Improvement
- Labour & Garden Execution

### Hardscape / Site Features

Examples include:

- Pathway
- Rock Garden
- Water Feature
- Seating Area
- Parking Landscape
- Irrigation System
- Landscape Lighting
- Children's Area

This allows the generated estimate to remain aligned with the selected project requirements.

---

## 📸 AI Site Visualization

Greenscape can generate a visual representation of the proposed landscape using the uploaded site photograph.

The visualization workflow is designed around a **plant-only modification approach**:

- Preserve the original site photograph
- Add plants from the generated Plant BOQ
- Maintain the existing site context
- Avoid unnecessary structural changes
- Avoid adding unrelated landscape elements

This provides a visual concept of how the selected planting scheme could appear on the site.

---

## 📊 Excel BOQ & Quotation Export

Greenscape can generate downloadable Excel-based project outputs.

These can be used for:

- BOQ preparation
- Cost estimation
- Quotation preparation
- Client discussions
- Project documentation

The generated spreadsheet helps reduce repetitive manual calculations.

---

# 🔄 Application Workflow

```text
        ┌──────────────────────┐
        │   Project Details    │
        └──────────┬───────────┘
                   │
                   ▼
        ┌──────────────────────┐
        │   Site / Soil Image  │
        └──────────┬───────────┘
                   │
                   ▼
        ┌──────────────────────┐
        │  Select Project Scope│
        └──────────┬───────────┘
                   │
                   ▼
        ┌──────────────────────┐
        │    AI Analysis       │
        └──────────┬───────────┘
                   │
          ┌────────┴─────────┐
          ▼                  ▼
 ┌─────────────────┐  ┌─────────────────┐
 │ Plant Recommend.│  │ Landscape Plan  │
 └────────┬────────┘  └────────┬────────┘
          │                    │
          └─────────┬──────────┘
                    ▼
          ┌───────────────────┐
          │     Plant BOQ     │
          └─────────┬─────────┘
                    │
             ┌──────┴───────┐
             ▼              ▼
      ┌─────────────┐ ┌───────────────┐
      │ Cost Estimate│ │ Site Visual   │
      └──────┬──────┘ └───────┬───────┘
             │                │
             └───────┬────────┘
                     ▼
          ┌────────────────────┐
          │ Excel / Quotation  │
          │      Export        │
          └────────────────────┘
````

---

# 🛠️ Technology Stack

| Technology                       | Purpose                                    |
| -------------------------------- | ------------------------------------------ |
| **Python**                       | Core application logic                     |
| **Streamlit**                    | Web application framework                  |
| **Google Gemini**                | AI-powered landscape analysis and planning |
| **Cloudflare Workers AI / FLUX** | AI site visualization                      |
| **Pandas**                       | Plant data processing                      |
| **Pillow**                       | Image processing                           |
| **OpenPyXL**                     | Excel BOQ / quotation generation           |
| **Requests**                     | API communication                          |
| **python-dotenv**                | Local environment configuration            |
| **CSV**                          | Plant database                             |

---

# 🏗️ Project Architecture

```text
Greenscape
│
├── app.py
│   └── Main Streamlit application
│
├── plants.csv
│   └── Plant database and pricing information
│
├── requirements.txt
│   └── Python dependencies
│
├── .gitignore
│   └── Protects environment files and local configuration
│
└── README.md
    └── Project documentation
```

---

# 📂 Dataset

The project uses a structured `plants.csv` database.

The database is used to support plant recommendation and estimation workflows.

Typical information includes:

```text
Plant Name
Category
Price
```

The application reads this database dynamically and uses it when generating the Plant BOQ.

---

# 🔐 Environment Variables

API credentials are **not stored inside the source code**.

For local development, create a `.env` file:

```env
GEMINI_API_KEY=your_gemini_api_key
CLOUDFLARE_API_TOKEN=your_cloudflare_api_token
CLOUDFLARE_ACCOUNT_ID=your_cloudflare_account_id
```

### Streamlit Cloud

For deployment, configure the same values through **Streamlit Secrets**:

```toml
GEMINI_API_KEY = "your_gemini_api_key"
CLOUDFLARE_API_TOKEN = "your_cloudflare_api_token"
CLOUDFLARE_ACCOUNT_ID = "your_cloudflare_account_id"
```

⚠️ **Never commit `.env` or API keys to GitHub.**

---

# 💻 Local Installation

## 1. Clone the repository

```bash
git clone https://github.com/Rishab295/greenscape-landscape-planner.git
```

## 2. Open the project

```bash
cd greenscape-landscape-planner
```

## 3. Install dependencies

```bash
pip install -r requirements.txt
```

## 4. Configure API keys

Create a `.env` file:

```env
GEMINI_API_KEY=your_gemini_api_key
CLOUDFLARE_API_TOKEN=your_cloudflare_api_token
CLOUDFLARE_ACCOUNT_ID=your_cloudflare_account_id
```

## 5. Run the application

```bash
streamlit run app.py
```

The application will open in your browser.

---

# ☁️ Deployment

Greenscape is deployed using:

**Streamlit Community Cloud**

Deployment workflow:

```text
GitHub Repository
        │
        ▼
Streamlit Community Cloud
        │
        ▼
Streamlit Application
        │
        ├── Gemini API
        │
        └── Cloudflare AI
```

Every update pushed to the GitHub repository can be deployed to the Streamlit application.

---

# 📸 Screenshots

> Add screenshots of the application here to make the GitHub repository more visually attractive.

Recommended screenshots:

### 1. Main Dashboard

```text
docs/screenshots/dashboard.png
```

### 2. Project Planning

```text
docs/screenshots/planning.png
```

### 3. Plant BOQ

```text
docs/screenshots/plant-boq.png
```

### 4. Cost Estimation

```text
docs/screenshots/quotation.png
```

### 5. AI Site Visualization

```text
docs/screenshots/visualization.png
```

Example:

```markdown
![Greenscape Dashboard](docs/screenshots/dashboard.png)
```

---

# 📈 Future Improvements

Potential future enhancements include:

* 📐 Automatic landscape area measurement
* 🗺️ Interactive site layout generation
* 🌦️ Weather and climate-based plant recommendations
* 🌱 Advanced soil analysis
* 📍 Location-based plant suitability
* 💧 Automated irrigation calculation
* 📊 Advanced project analytics
* 🧾 Professional PDF quotation generation
* 🏡 2D landscape layout generation
* 🤖 RAG-based landscaping knowledge assistant
* 📱 Mobile-optimized interface
* 👥 Multi-user project management
* 💾 Project history and database storage

---

# 🎯 Real-World Problem Solved

Traditional landscape planning can require significant manual effort for:

* Plant selection
* Quantity calculation
* BOQ preparation
* Cost estimation
* Client visualization
* Quotation preparation

Greenscape brings these tasks into a single AI-assisted workflow.

```text
Manual Landscape Planning
          ↓
Multiple Tools
          ↓
Manual Calculations
          ↓
Manual BOQ
          ↓
Manual Quotation
          ↓
Manual Visualization
```

### Greenscape

```text
Project Requirements
          ↓
AI Landscape Analysis
          ↓
Plant Recommendations
          ↓
Plant BOQ
          ↓
Cost Estimation
          ↓
Site Visualization
          ↓
Excel / Quotation
```

---

# 👨‍💻 Developer

**Rishab Das**

MSc Data Science

### GitHub

[![GitHub](https://img.shields.io/badge/GitHub-Rishab295-black?style=for-the-badge\&logo=github)](https://github.com/Rishab295)

### Live Application

[![Streamlit](https://img.shields.io/badge/Live%20Demo-Streamlit-red?style=for-the-badge\&logo=streamlit)](https://greenscape-landscape-planner-kwhp4zqojnhweior5ubnux.streamlit.app/)

---

# 📄 License

This project currently does not include an open-source license.

---

<p align="center">

### 🌿 Greenscape

**AI Landscape Planning & Estimation Assistant**

Built with Python • Streamlit • Gemini • Cloudflare AI

</p>
```


