# European EdTech Ecosystem Survey (WeMap Form Specification)

This specification reverse-engineers the exact survey form structure that generated the **WeMap Clean Dataset** (`202507_dataset_WeMap_LIVE_202607281022_CLEAN.xlsx`).

---

## Form Metadata
* **Form Title:** European EdTech Ecosystem Directory & Map (WeMap Survey)
* **Form Description:** 
  > Welcome to the European EdTech Map (WeMap) initiative. Please fill out this survey to showcase your educational technology organization on the European EdTech Explorer platform. Your submission helps educators, institutions, policymakers, and investors discover European EdTech solutions.
* **Estimated Completion Time:** 7–10 minutes

---

## How Google Forms Handles "Other" Options & Spreadsheet Output

When a question has pre-defined checkboxes/radios plus an **"Other"** option, Google Forms handles them as follows:

1. **How it saves in the Excel / Google Sheets output:**
   * **For Checkboxes (multi-select):** Google Forms combines all selected choices into a **single comma-separated string** in that column. If a respondent selects `Business to Business (B2B)` and checks `Other` typing `Revenue share on book sales`, the cell in the spreadsheet will literally contain:
     ```
     Business to Business, Revenue share on book sales
     ```
     *(Notice the word "Other:" is automatically replaced by whatever the respondent typed into the box).*
   * **If the user selects a predefined option AND does NOT select "Other":** The text box next to "Other" remains inactive/empty, and only the checked pre-defined options are saved.
   * **If the user checks "Other" but leaves the text box blank:** Google Forms saves the literal text `"Other"` in that cell.
2. **Forcing/Requiring an explanation when "Other" is chosen:**
   * In standard Google Forms, native inline "Other" cannot be made conditionally compulsory within the same checkbox widget.
   * **The Solution:** A dedicated follow-up short-answer field is placed right below the question:
     > *"If you selected 'Other' above, please specify here:"*
     During data processing, if this follow-up field has a value, it is merged into the main column or reviewed by the pipeline.

---

## Output Column Mapping Overview

| Survey Question | Field Type | Target XLSX Column | Required? |
| :--- | :--- | :--- | :--- |
| *Form System Timestamp* | Automatic | `last_update_date` | Automatic |
| *Database Unique ID* | Reviewer/Pipeline | `org_id` | Automatic |
| *Status Flag* | Reviewer/Pipeline | `is_current` | Automatic |
| **Organization Public Name** | Short answer | `public_name` | **Required** |
| **Official Website URL** | Short answer | `website` | **Required** |
| **Organization Description** | Paragraph | `org_desc` | **Required** |
| **Member Association(s)** | Checkboxes (+ Other) | `member_association` | Optional |
| ↳ *Other Association Name* | Short answer | `member_association` (merged) | Conditional |
| **Primary City / Headquarters** | Short answer | `city` | **Required** |
| **Headquarters Country** | Dropdown | `country` | **Required** |
| **Countries where you sell/operate** | Checkboxes / Paragraph | `sell_countries` | **Required** |
| **Company Scale / Team Size (FTE)** | Multiple choice | `employees_number` | **Required** |
| **Profit / Legal Status** | Multiple choice | `profit_status` | Optional |
| **Business Model(s)** | Checkboxes (+ Other) | `business_model` | **Required** |
| ↳ *Other Business Model* | Short answer | `business_model` (merged) | Optional |
| **Target Market Segment(s)** | Checkboxes (+ Other) | `market_segments` | **Required** |
| ↳ *Other Market Segment* | Short answer | `market_segments` (merged) | Optional |
| **Target Audience / End User** | Checkboxes (+ Other) | `target_audience` | **Required** |
| ↳ *Other Target Audience* | Short answer | `target_audience` (merged) | Optional |
| **Offer solutions for physical settings?** | Multiple choice (Yes/No) | `edu_settings` | Optional |
| **Offer educational content / curricula?** | Multiple choice (Yes/No) | `edu_content` | Optional |
| **Supported Product Languages** | Checkboxes (+ Other) | `product_language` | Optional |
| **Flagship Product 1 — Name** | Short answer | `product_1_name` | Optional |
| **Flagship Product 1 — Description / URL** | Paragraph | `product_1_desc` | Optional |
| **Product 2 — Name** | Short answer | `product_2_name` | Optional |
| **Product 2 — Description / URL** | Paragraph | `product_2_desc` | Optional |
| **Company Logo Upload** | File upload (Image) | `WeMap_URL` | **Required** |
| **LinkedIn Company Page URL** | Short answer | `socials_linkedin` | Optional |
| **X / Twitter Profile URL** | Short answer | `socials_x` | Optional |
| **Instagram Profile URL** | Short answer | `socials_insta` | Optional |
| **Facebook Page URL** | Short answer | `facebook` | Optional |
| **Partnership Interests** | Checkboxes (+ Other) | `partnerships` | Optional |
| *Reviewer Verification* | Reviewer/Pipeline | `verified` | Pipeline |

---

## Detailed Section-by-Section Form Layout

### Section 1: Organization Overview
*(Styled with a full section header bar for aesthetic consistency with Section 2)*

#### Q1. Organization Public Name `[public_name]`
* **Type:** Short answer text
* **Required:** Yes
* **Help text:** Enter the official public trading name of your organization or startup.

#### Q2. Official Website URL `[website]`
* **Type:** Short answer text
* **Required:** Yes
* **Help text:** Example: `www.company.com` or `company.eu` *(plain text example, avoiding clickable dead links)*

#### Q3. Organization Description `[org_desc]`
* **Type:** Paragraph
* **Required:** Yes
* **Help text:** A concise overview (2–5 sentences) of your company's mission, pedagogical focus, and main solution.

#### Q4. EdTech Member Association(s) `[member_association]`
* **Type:** Checkboxes (+ "Other" text write-in)
* **Required:** No
* **Help text:** Select all national or European EdTech clusters/associations your organization belongs to. If not listed, select "Other".
* **Options:**
  - AFINEF (France)
  - bitmark Association
  - DOHE
  - Dutch EdTech (Netherlands)
  - edtechstation (Belgium)
  - EdTech Austria
  - EdTech Bulgaria
  - EdTech Denmark
  - EdTech Estonia
  - EdTech Finland
  - EdTech France
  - EdTech Ireland
  - EdTech Italia
  - EdTech Lithuania
  - EdTech Poland
  - EdTech Slovenia
  - EdTech Ukraine
  - EdTech-Verband e.V. (Germany)
  - EduCloud Alliance
  - EduTech Cluster Spain
  - EduTech Hungary
  - Education:NEXT
  - EDUvation
  - ETKİM Türkiye
  - ICT Norway
  - LearnTech Hub
  - SEK Lab
  - SuperCharger Ventures
  - Swedish EdTech Industry
  - Swiss EdTech Collider
  - xEdu
  - I'm not part of any EdTech association
  - Other...

#### Q4b. If you selected 'Other' for Member Association, please state its full name:
* **Type:** Short answer text
* **Required:** Conditional (required if "Other" was ticked)
* **Help text:** Enter the official name of the EdTech organization, hub, or regional cluster.

---

### Section 2: Headquarters & Geographic Reach

#### Q5. Headquarters City `[city]`
* **Type:** Short answer text
* **Required:** Yes
* **Help text:** The city where your primary headquarters or registered office is located.

#### Q6. Headquarters Country `[country]`
* **Type:** Dropdown
* **Required:** Yes
* **Help text:** Country where your headquarters is located.
* **Options:**
  Austria, Belgium, Bulgaria, Croatia, Cyprus, Czech Republic, Denmark, Estonia, Finland, France, Germany, Greece, Hungary, Iceland, Ireland, Italy, Latvia, Lithuania, Luxembourg, Malta, Netherlands, Norway, Poland, Portugal, Romania, Slovakia, Slovenia, Spain, Sweden, Switzerland, Turkey, Ukraine, United Kingdom, Other.

#### Q7. Countries of Operation / Sales `[sell_countries]`
* **Type:** Paragraph
* **Required:** Yes
* **Help text:** List the primary countries where your solutions are sold or actively used (e.g. `Germany, France, United Kingdom, Global`).

---

### Section 3: Scale, Business Model & Audience

#### Q8. Company Scale / Team Size (FTE) `[employees_number]`
* **Type:** Multiple choice
* **Required:** Yes
* **Help text:** Select the standardized EU SME category representing your full-time equivalent (FTE) team size:
* **Options:**
  - **Solo entrepreneur (1 employee/Self)**
  - **Micro (1-9 FTE employees)**
  - **Small (10-49 FTE employees)**
  - **Medium (50-250 FTE employees)**
  - **Large (250+ FTE employees)**

#### Q9. Legal / Profit Status `[profit_status]`
* **Type:** Multiple choice (+ "Other")
* **Required:** No
* **Help text:** *Academic spin-offs are independent companies founded to commercialize research, intellectual property, or pedagogical tools developed within a university or higher education research institute.*
* **Options:**
  - For-profit company
  - Non-profit organisation
  - Public institution
  - University / Academic spin-off
  - Other...

#### Q10. Business Model(s) `[business_model]`
* **Type:** Checkboxes (+ "Other" text write-in)
* **Required:** Yes
* **Options:**
  - Business to Business (B2B)
  - Business to Schools / Institutions (B2S)
  - Business to Consumer (B2C)
  - Business to Government (B2G)
  - Other...

#### Q10b. If you selected 'Other' for Business Model, please specify:
* **Type:** Short answer text
* **Required:** No

#### Q11. Market Segment(s) `[market_segments]`
* **Type:** Checkboxes (+ "Other" text write-in)
* **Required:** Yes
* **Help text:** Select all sectors your solutions serve.
* **Options:**
  - Early childhood development (0-5 years)
  - Primary Education
  - Secondary Education
  - Higher Education
  - Vocational Education and Training (VET)
  - Corporate Training
  - Lifelong learning / Adult education
  - Special Education (SEN)
  - Other...

#### Q11b. If you selected 'Other' for Market Segment, please specify:
* **Type:** Short answer text
* **Required:** No

#### Q12. Target Audience / End User Interface `[target_audience]`
* **Type:** Checkboxes (+ "Other" text write-in)
* **Required:** Yes
* **Options:**
  - Student interfacing
  - Teacher interfacing
  - System/Administrative
  - Other...

#### Q12b. If you selected 'Other' for Target Audience, please specify:
* **Type:** Short answer text
* **Required:** No

#### Q13. Physical Educational Settings `[edu_settings]`
* **Type:** Multiple choice
* **Required:** No
* **Question:** Does your solution serve physical, in-person educational settings (hardware, classroom tools, on-premise)?
* **Options:**
  - Yes
  - No

#### Q14. Educational Content / Curriculum `[edu_content]`
* **Type:** Multiple choice
* **Required:** No
* **Question:** Does your company directly produce or license instructional content/curriculum materials (as opposed to pure software/infrastructure)?
* **Options:**
  - Yes
  - No

---

### Section 4: Products & Offerings

#### Q15. Supported Product Languages `[product_language]`
* **Type:** Checkboxes (+ "Other" text write-in)
* **Required:** No
* **Help text:** Languages in which your product / user interface is available.
* **Options:**
  English, French, German, Spanish, Italian, Dutch, Polish, Portuguese, Swedish, Finnish, Danish, Norwegian, Ukrainian, Other...

#### Q16. Primary Product 1 — Name `[product_1_name]`
* **Type:** Short answer text
* **Required:** No

#### Q17. Primary Product 1 — Description / Link `[product_1_desc]`
* **Type:** Paragraph
* **Required:** No

#### Q18. Product 2 — Name `[product_2_name]`
* **Type:** Short answer text
* **Required:** No

#### Q19. Product 2 — Description / Link `[product_2_desc]`
* **Type:** Paragraph
* **Required:** No

---

### Section 5: Logo & Online Profiles

#### Q20. Organization Logo `[WeMap_URL]`
* **Type:** File upload (Google Drive file upload item)
* **Required:** Yes
* **Help text:** Upload a clean PNG, JPG, or SVG of your organization logo (square or horizontal transparent background preferred).
* **Spreadsheet Output Note:** Google Forms automatically places the uploaded file into the form owner's Google Drive folder and writes the direct `https://drive.google.com/uc?export=view&id=...` link into the `WeMap_URL` column.

#### Q21. LinkedIn Page URL `[socials_linkedin]`
* **Type:** Short answer text
* **Required:** No

#### Q22. X / Twitter Profile URL `[socials_x]`
* **Type:** Short answer text
* **Required:** No

#### Q23. Instagram Profile URL `[socials_insta]`
* **Type:** Short answer text
* **Required:** No

#### Q24. Facebook Page URL `[facebook]`
* **Type:** Short answer text
* **Required:** No

---

### Section 6: Ecosystem Collaboration

#### Q25. Partnerships & Collaboration Interests `[partnerships]`
* **Type:** Checkboxes (+ "Other" text write-in)
* **Required:** No
* **Help text:** What forms of ecosystem partnerships is your organization open to?
* **Options:**
  - Collaborations between EdTech organisations
  - Industry-academia partnerships
  - Public-private partnerships
  - Project consortium building
  - Other...
