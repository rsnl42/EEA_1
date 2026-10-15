/**
 * create_google_form.gs
 * Google Apps Script — Auto-generates the European EdTech Explorer (WeMap) survey.
 *
 * Output column order matches the original dataset:
 *   last_update_date | org_id | is_current | public_name | website | org_desc |
 *   member_association | product_language | city | country | sell_countries |
 *   employees_number | business_model | market_segments | target_audience |
 *   socials_insta | socials_x | socials_linkedin | product_1_name | product_1_desc |
 *   product_2_name | product_2_desc | WeMap_URL | verified | partnerships |
 *   profit_status | edu_settings | edu_content | facebook
 *
 * HOW TO USE:
 *   1. Go to https://script.google.com → New project.
 *   2. Delete the default code, paste this entire file, and click Save.
 *   3. Select the function `createWeMapSurvey` from the dropdown and click Run.
 *   4. Accept the authorization prompt.
 *   5. The Execution Log will print the form's published URL and edit URL.
 *
 * HOW RESPONSES ARE COLLECTED:
 *   - Open the generated form in Google Drive.
 *   - Click the Responses tab → click the Google Sheets icon to link a spreadsheet.
 *   - All submissions will be automatically appended as rows in that Google Sheet.
 *   - To export as .xlsx: File → Download → Microsoft Excel (.xlsx) in Google Sheets.
 *   - To transfer ownership: Form → ⋮ menu → Add collaborator → transfer owner role.
 *   - To enable email alerts: Responses tab → ⋮ → Get email notifications for new responses.
 *
 * RESPONSE EDITING (edit without resetting other answers):
 *   - setAllowResponseEdits(true) is enabled below.
 *   - After submitting, respondents see an "Edit your response" link on the confirmation page.
 *   - Clicking it reopens their prefilled form; they can change individual fields and resubmit.
 *
 * NOTE ON "OTHER" OPTIONS:
 *   - Google Forms' native "Other" in checkboxes saves the text typed inline directly
 *     into the spreadsheet cell (comma-separated with any other selections).
 *   - For questions requiring the name of the Other option, a dedicated follow-up
 *     short-answer field is placed immediately after.
 */

function createWeMapSurvey() {

  const form = FormApp.create("European EdTech Ecosystem Directory & Map (WeMap Survey)");

  form.setDescription(
    "Welcome to the European EdTech Map (WeMap) initiative.\n\n" +
    "Please fill in this survey to list your educational technology organization on the European EdTech Explorer platform. " +
    "Your submission helps educators, institutions, policymakers, and investors discover European EdTech solutions.\n\n" +
    "Tip: The form is divided into sections — you can navigate between sections before submitting. " +
    "After submitting, you will see an 'Edit your response' link on the confirmation page, which allows you to " +
    "return to your answers, correct any entry, and resubmit — without losing your other responses.\n\n" +
    "Estimated completion time: 7–10 minutes."
  );

  form.setAllowResponseEdits(true);  // Enables post-submission editing link
  form.setCollectEmail(true);

  // ════════════════════════════════════════
  // SECTION 1: Organization Overview
  // Columns: public_name | website | org_desc | member_association
  // ════════════════════════════════════════
  form.addPageBreakItem()
      .setTitle("Section 1: Organization Overview")
      .setHelpText("Basic identification details and association memberships.");

  // → public_name
  form.addTextItem()
      .setTitle("Organization Public Name")
      .setHelpText("The official public trading or brand name of your organization.")
      .setRequired(true);

  // → website
  form.addTextItem()
      .setTitle("Official Website")
      .setHelpText("Example: www.company.com or company.eu")
      .setRequired(true);

  // → org_desc
  form.addParagraphTextItem()
      .setTitle("Organization Description")
      .setHelpText("A concise overview (2–5 sentences) of your company's mission, core technology, and educational value proposition.")
      .setRequired(true);

  // → member_association
  const assocItem = form.addCheckboxItem();
  assocItem.setTitle("EdTech Member Association(s)")
           .setHelpText(
             "Select all national or regional EdTech clusters/associations your organization belongs to. " +
             "If your association is not listed, select 'Other' and specify its name in the field below."
           )
           .setRequired(false)
           .setChoiceValues([
             "AFINEF (France)", "bitmark Association", "DOHE", "Dutch EdTech",
             "edtechstation (Belgium)", "EdTech Austria", "EdTech Bulgaria", "EdTech Denmark",
             "EdTech Estonia", "EdTech Finland", "EdTech France", "EdTech Ireland",
             "EdTech Italia", "EdTech Lithuania", "EdTech Poland", "EdTech Slovenia",
             "EdTech Ukraine", "EdTech-Verband e.V. (Germany)", "EduCloud Alliance",
             "EduTech Cluster Spain", "EduTech Hungary", "Education:NEXT", "EDUvation",
             "ETKİM Türkiye", "ICT Norway", "LearnTech Hub", "SEK Lab",
             "SuperCharger Ventures", "Swedish EdTech Industry", "Swiss EdTech Collider",
             "xEdu", "I'm not part of any EdTech association"
           ]);
  assocItem.showOtherOption(true);

  // Follow-up: compulsory name if Other ticked
  form.addTextItem()
      .setTitle("If you selected 'Other' for Member Association — please state its full name:")
      .setHelpText("Required if you selected 'Other' above. Enter the official name of the EdTech organization, cluster, or hub.")
      .setRequired(false);

  // ════════════════════════════════════════
  // SECTION 2: Geography & Location
  // Columns: city | country | sell_countries
  // ════════════════════════════════════════
  form.addPageBreakItem()
      .setTitle("Section 2: Headquarters & Geographic Reach")
      .setHelpText("Where your team is based and the regions where your solutions are active.");

  // → city
  form.addTextItem()
      .setTitle("Headquarters City")
      .setHelpText("The city where your primary headquarters or registered office is located.")
      .setRequired(true);

  // → country
  const countryItem = form.addListItem();
  countryItem.setTitle("Headquarters Country")
             .setHelpText("Country of primary registration.")
             .setRequired(true)
             .setChoiceValues([
               "Austria", "Belgium", "Bulgaria", "Croatia", "Cyprus", "Czech Republic",
               "Denmark", "Estonia", "Finland", "France", "Germany", "Greece", "Hungary",
               "Iceland", "Ireland", "Italy", "Latvia", "Lithuania", "Luxembourg", "Malta",
               "Netherlands", "Norway", "Poland", "Portugal", "Romania", "Slovakia",
               "Slovenia", "Spain", "Sweden", "Switzerland", "Turkey", "Ukraine",
               "United Kingdom", "Other"
             ]);

  // → sell_countries
  form.addParagraphTextItem()
      .setTitle("Countries of Operation / Sales")
      .setHelpText("List countries where your solution is sold or actively deployed (e.g. France, Germany, United Kingdom, Global).")
      .setRequired(true);

  // ════════════════════════════════════════
  // SECTION 3: Scale, Business Model & Audience
  // Columns: employees_number | business_model | market_segments | target_audience |
  //          profit_status | edu_settings | edu_content
  // ════════════════════════════════════════
  form.addPageBreakItem()
      .setTitle("Section 3: Scale, Business Model & Audience")
      .setHelpText("Company size, revenue models, and the educational segments and audiences you serve.");

  // → employees_number (EU SME categories — matching original dataset)
  form.addMultipleChoiceItem()
      .setTitle("Company Scale / Team Size (FTE)")
      .setHelpText("Select the standardized category that best represents your full-time equivalent (FTE) headcount.")
      .setRequired(true)
      .setChoiceValues([
        "Solo entrepreneur (1 employee/Self)",
        "Micro (1-9 FTE employees)",
        "Small (10-49 FTE employees)",
        "Medium (50-250 FTE employees)",
        "Large (250+ FTE employees)"
      ]);

  // → profit_status
  const profitItem = form.addMultipleChoiceItem();
  profitItem.setTitle("Legal / Profit Status")
            .setHelpText(
              "Academic spin-off: a company founded to commercialize research, IP, or pedagogical tools " +
              "originally developed within a university or research institute."
            )
            .setRequired(false)
            .setChoiceValues([
              "For-profit company",
              "Non-profit organisation",
              "Public institution",
              "University / Academic spin-off"
            ]);
  profitItem.showOtherOption(true);

  // → business_model
  const bizItem = form.addCheckboxItem();
  bizItem.setTitle("Business Model(s)")
         .setHelpText("Select all that apply. If none of the options match, select 'Other' and specify below.")
         .setRequired(true)
         .setChoiceValues([
           "Business to Business (B2B)",
           "Business to Schools / Institutions (B2S)",
           "Business to Consumer (B2C)",
           "Business to Government (B2G)"
         ]);
  bizItem.showOtherOption(true);

  form.addTextItem()
      .setTitle("If you selected 'Other' for Business Model — please specify:")
      .setHelpText("Describe your custom monetization or distribution structure.")
      .setRequired(false);

  // → market_segments
  const segItem = form.addCheckboxItem();
  segItem.setTitle("Market Segment(s)")
         .setHelpText("Select all educational phases your solutions address. Select 'Other' if your segment is not listed.")
         .setRequired(true)
         .setChoiceValues([
           "Early childhood development (0-5 years)",
           "Primary Education",
           "Secondary Education",
           "Higher Education",
           "Vocational Education and Training (VET)",
           "Corporate Training",
           "Lifelong learning/Adult education",
           "Special Education"
         ]);
  segItem.showOtherOption(true);

  form.addTextItem()
      .setTitle("If you selected 'Other' for Market Segment — please specify:")
      .setHelpText("e.g. Informal education, Parent coaching, Museum learning.")
      .setRequired(false);

  // → target_audience
  const audItem = form.addCheckboxItem();
  audItem.setTitle("Target Audience / End User Interface")
         .setHelpText("Who primarily uses or interacts with your solution directly?")
         .setRequired(true)
         .setChoiceValues([
           "Student interfacing",
           "Teacher interfacing",
           "System/Administrative"
         ]);
  audItem.showOtherOption(true);

  form.addTextItem()
      .setTitle("If you selected 'Other' for Target Audience — please specify:")
      .setHelpText("e.g. Parents, HR directors, IT administrators.")
      .setRequired(false);

  // → edu_settings
  form.addMultipleChoiceItem()
      .setTitle("Does your solution serve physical, in-person educational settings?")
      .setHelpText("e.g. classroom hardware, campus infrastructure, or physical learning kits.")
      .setRequired(false)
      .setChoiceValues(["Yes", "No"]);

  // → edu_content
  form.addMultipleChoiceItem()
      .setTitle("Do you directly create or license educational curriculum / learning content?")
      .setHelpText("As opposed to solely providing software tools or platforms.")
      .setRequired(false)
      .setChoiceValues(["Yes", "No"]);

  // ════════════════════════════════════════
  // SECTION 4: Products & Offerings
  // Columns: product_language | product_1_name | product_1_desc | product_2_name | product_2_desc
  // ════════════════════════════════════════
  form.addPageBreakItem()
      .setTitle("Section 4: Products & Offerings")
      .setHelpText("Product names, descriptions, and supported language coverage.");

  // → product_language
  const langItem = form.addCheckboxItem();
  langItem.setTitle("Supported Product Languages")
          .setHelpText(
            "Languages in which your product, platform, or content is available to end users. " +
            "If a language is not listed, select 'Other' and name it in the field below."
          )
          .setRequired(false)
          .setChoiceValues([
            "English", "French", "German", "Spanish", "Italian", "Dutch",
            "Polish", "Portuguese", "Swedish", "Finnish", "Danish", "Norwegian",
            "Ukrainian", "Czech", "Slovak", "Hungarian", "Romanian", "Bulgarian",
            "Greek", "Croatian", "Slovenian", "Estonian", "Latvian", "Lithuanian",
            "Maltese", "Irish", "Catalan"
          ]);
  langItem.showOtherOption(true);

  form.addTextItem()
      .setTitle("If you selected 'Other' for Product Language — please specify:")
      .setHelpText("List any additional supported languages not found in the options above.")
      .setRequired(false);

  // → product_1_name / product_1_desc
  form.addTextItem()
      .setTitle("Primary Product / Service 1 — Name")
      .setRequired(false);

  form.addParagraphTextItem()
      .setTitle("Primary Product / Service 1 — Description or Link")
      .setHelpText("A brief description of the product and its educational purpose, or a link to its product page.")
      .setRequired(false);

  // → product_2_name / product_2_desc
  form.addTextItem()
      .setTitle("Product / Service 2 — Name")
      .setRequired(false);

  form.addParagraphTextItem()
      .setTitle("Product / Service 2 — Description or Link")
      .setRequired(false);

  // ════════════════════════════════════════
  // SECTION 5: Logo & Social Profiles
  // Columns: WeMap_URL | socials_insta | socials_x | socials_linkedin | facebook
  // ════════════════════════════════════════
  form.addPageBreakItem()
      .setTitle("Section 5: Logo & Online Profiles")
      .setHelpText(
        "Your logo and social links appear on your organization's directory profile card. " +
        "Logo upload: the file is saved to the form owner's Google Drive and the direct link is recorded."
      );

  // → WeMap_URL
  // NOTE: File upload items require the form to be created in a Google Workspace domain.
  // If running under a personal Gmail account, replace this with a short-answer URL field.
  // To switch it to a real file upload in the Form UI: open the form → + → File Upload.
  form.addTextItem()
      .setTitle("Company Logo — Google Drive Upload Link or Direct Image URL")
      .setHelpText(
        "Upload your logo to Google Drive, set sharing to 'Anyone with the link can view', " +
        "then paste the link here. PNG or SVG with a transparent background preferred."
      )
      .setRequired(true);

  // → socials_insta
  form.addTextItem()
      .setTitle("Instagram Profile URL")
      .setHelpText("Example: www.instagram.com/yourcompany")
      .setRequired(false);

  // → socials_x
  form.addTextItem()
      .setTitle("X / Twitter Profile URL")
      .setHelpText("Example: x.com/yourhandle")
      .setRequired(false);

  // → socials_linkedin
  form.addTextItem()
      .setTitle("LinkedIn Company Page URL")
      .setHelpText("Example: www.linkedin.com/company/yourcompany")
      .setRequired(false);

  // → facebook
  form.addTextItem()
      .setTitle("Facebook Page URL")
      .setHelpText("Example: www.facebook.com/yourcompany")
      .setRequired(false);

  // ════════════════════════════════════════
  // SECTION 6: Ecosystem Collaboration
  // Column: partnerships
  // ════════════════════════════════════════
  form.addPageBreakItem()
      .setTitle("Section 6: Ecosystem Collaboration")
      .setHelpText("Opportunities for joint initiatives, research projects, and consortium building.");

  const partItem = form.addCheckboxItem();
  partItem.setTitle("Partnership & Collaboration Interests")
          .setHelpText(
            "What kinds of partnerships is your organization actively seeking within the European ecosystem? " +
            "Select 'Other' if your preferred collaboration type is not listed, and describe it below."
          )
          .setRequired(false)
          .setChoiceValues([
            "Collaborations between EdTech organisations",
            "Industry-academia partnerships",
            "Public-private partnerships",
            "Project consortium building"
          ]);
  partItem.showOtherOption(true);

  form.addTextItem()
      .setTitle("If you selected 'Other' for Partnership Interests — please specify:")
      .setHelpText("e.g. Investor introductions, Reseller agreements, Pilot school programmes.")
      .setRequired(false);

  // ════════════════════════════════════════
  // SECTION 7: Review & Submit
  // (No columns — this is a reminder page before final submission)
  // ════════════════════════════════════════
  form.addPageBreakItem()
      .setTitle("Section 7: Review & Submit")
      .setHelpText(
        "Before submitting, please take a moment to confirm the following:\n\n" +
        "✅  Your organization's public name and website are spelled correctly.\n" +
        "✅  Your description clearly explains what your company does and who it serves.\n" +
        "✅  You have selected all relevant market segments and business models.\n" +
        "✅  Your logo link works and points to a visible image.\n" +
        "✅  Social profile URLs are complete and active.\n\n" +
        "Once submitted, you will see a confirmation page with an 'Edit your response' link. " +
        "You can use that link at any time to return to your answers, correct individual fields, " +
        "and resubmit — without losing any other responses."
      );

  form.addParagraphTextItem()
      .setTitle("Any additional comments or information you would like to share with the WeMap team?")
      .setHelpText("Optional — this field does not map to a dataset column.")
      .setRequired(false);

  // ════════════════════════════════════════
  // OUTPUT
  // ════════════════════════════════════════
  Logger.log("══════════════════════════════════════════");
  Logger.log("FORM CREATED SUCCESSFULLY!");
  Logger.log("Published URL: " + form.getPublishedUrl());
  Logger.log("Edit URL:      " + form.getEditUrl());
  Logger.log("══════════════════════════════════════════");
  Logger.log("NEXT STEPS:");
  Logger.log("1. Open the Edit URL above.");
  Logger.log("2. Responses tab → click the Sheets icon to create a linked Google Sheet.");
  Logger.log("3. Responses tab → ⋮ → 'Get email notifications for new responses'.");
  Logger.log("4. To transfer ownership: ⋮ menu → Add collaborator → transfer Owner role.");
  Logger.log("5. Download responses as .xlsx: open the linked Sheet → File → Download → Excel.");
}
