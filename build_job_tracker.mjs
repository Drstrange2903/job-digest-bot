import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "outputs/job_tracker";
await fs.mkdir(outputDir, { recursive: true });

const workbook = Workbook.create();
const overview = workbook.worksheets.add("Overview");
const tracker = workbook.worksheets.add("Job Tracker");
const templates = workbook.worksheets.add("Templates");

const jobs = [
  ["To Review", 1, "BayOne Solutions", "Artificial Intelligence Engineer", "AI/ML", "India", "Remote", "", "LinkedIn", "https://www.linkedin.com/jobs/view/4416388565/", "", "Easy Apply; actively reviewing applicants. Good AI engineer title fit.", "Open job and review Easy Apply form", "", ""],
  ["To Review", 2, "Precision AQ", "AI/ML Engineer II", "AI/ML", "Bengaluru", "Remote", "", "LinkedIn", "https://www.linkedin.com/jobs/view/4417508818/", "", "Easy Apply; strong title and remote mode.", "Open job and review Easy Apply form", "", ""],
  ["To Review", 3, "Chubb", "AI/ML Engineer", "AI/ML", "Bengaluru", "Hybrid", "", "LinkedIn", "https://www.linkedin.com/jobs/view/4404058101/", "", "Easy Apply; actively reviewing applicants.", "Open job and review Easy Apply form", "", ""],
  ["To Review", 4, "Teradata", "AI Engineer", "AI/ML", "Pune/Pimpri-Chinchwad", "Hybrid", "", "LinkedIn", "https://www.linkedin.com/jobs/view/4370700303/", "", "Easy Apply; actively reviewing applicants. Good Pune option.", "Open job and review Easy Apply form", "", ""],
  ["To Review", 5, "HackerRank", "Machine Learning Engineer, Integrity", "Machine Learning", "Bengaluru East", "Hybrid", "", "LinkedIn", "https://www.linkedin.com/jobs/view/4398830613/", "", "Strong ML engineering brand; likely competitive.", "Review JD and tailor resume keywords", "", ""],
  ["To Review", 6, "HackerRank", "Machine Learning Engineer, Chakra", "Machine Learning", "Bengaluru East", "Hybrid", "", "LinkedIn", "https://www.linkedin.com/jobs/view/4398852823/", "", "Strong ML engineering opportunity.", "Review JD and tailor resume keywords", "", ""],
  ["To Review", 7, "5paisa", "AI/ML Engineer", "AI/ML", "Bengaluru", "On-site", "", "LinkedIn", "https://www.linkedin.com/jobs/view/4417130707/", "", "Easy Apply; fintech AI/ML role.", "Open Easy Apply form", "", ""],
  ["To Review", 8, "GoodScore", "AI/ML Engineer", "AI/ML", "Bengaluru", "On-site", "", "LinkedIn", "https://www.linkedin.com/jobs/view/4412886892/", "", "Easy Apply; actively reviewing applicants.", "Open Easy Apply form", "", ""],
  ["To Review", 9, "Helo.ai by VivaConnect", "Artificial Intelligence Engineer", "AI/ML", "Mumbai", "On-site", "", "LinkedIn", "https://www.linkedin.com/jobs/view/4410513025/", "", "Easy Apply; actively reviewing applicants. Good Mumbai option.", "Open Easy Apply form", "", ""],
  ["To Review", 10, "Neuron7.ai", "Software Engineer - AI and Python", "AI/Python", "Bengaluru", "On-site", "", "LinkedIn", "https://www.linkedin.com/jobs/view/4414458690/", "", "AI + Python role, likely product/startup-style.", "Review JD and apply if experience fit is good", "", ""],
  ["To Review", 11, "Verveo", "Generative AI Engineer", "Generative AI", "Bengaluru", "Not specified", "2-3 yrs", "LinkedIn/Public", "https://in.linkedin.com/jobs/view/generative-ai-engineer-at-verveo-4395975109", "hr@verveo.com", "Strong RAG/vector DB/LangChain/Python/Docker match.", "Email resume and apply on LinkedIn", "", ""],
  ["To Review", 12, "APPIT Software Solutions", "Generative AI Engineer", "Generative AI", "Hyderabad", "On-site", "2-4 yrs", "Company Careers", "https://www.appitsoftware.com/careers/generative-ai-engineer-hyderabad", "", "GenAI, RAG, vector DBs, LangChain/OpenAI, FastAPI, Hugging Face.", "Apply on company careers page", "", ""],
  ["To Review", 13, "NeuLeap", "Generative AI Engineer", "Generative AI", "Pune", "Full-time", "2-4 yrs", "Company Careers", "https://neuleap.ai/careers/5", "admin@neuleap.ai", "RAG, LangChain/LangGraph/LlamaIndex, Python, REST APIs, Docker/K8s.", "Apply and email resume", "", ""],
  ["To Review", 14, "coretek labs", "AI Engineer - GenAI / Agentic AI", "Agentic AI", "Hyderabad", "Not specified", "2-4 yrs", "Foundit", "https://www.foundit.in/job/ai-engineer-genai-agentic-ai-coretek-labs-hyderabad-secunderabad-telangana-48809971", "", "AI agents, RAG, LLMs, AWS/GCP, LangChain/LangGraph, REST APIs.", "Apply via Foundit", "", ""],
  ["To Review", 15, "Impetus Technologies", "Generative AI Developer / Engineer", "Generative AI", "Bangalore / Hyderabad / Pune / Indore", "Not specified", "2-4 yrs", "Talent.com", "https://in.talent.com/view?id=7f2543a9f40b", "", "Python, cloud, LLMs, RAG frameworks.", "Verify active listing and apply", "", ""],
  ["To Review", 16, "PwC", "Machine Learning Engineer - Senior Associate", "Machine Learning", "Bengaluru", "Not specified", "2-4 yrs", "LinkedIn", "https://in.linkedin.com/jobs/view/machine-learning-engineer-senior-associate-at-pwc-acceleration-center-india-4353535055", "", "GenAI, RAG, agentic AI, LangChain/LangGraph/AutoGen/CrewAI, Python, MLOps.", "Review requirements and apply", "", ""],
  ["Watch", 17, "Lucidspire", "Data Engineer", "Data Engineering", "Bengaluru", "Not specified", "2-5 yrs", "LinkedIn/Public", "https://in.linkedin.com/company/lucidspire", "divya.s@lucidspire.com", "SQL, ETL, Python basics, Power BI/Azure Data Factory exposure.", "Email resume if Data Engineer track is acceptable", "", ""],
  ["Watch", 18, "phData", "Data Engineer", "Data Engineering", "Bengaluru", "Not specified", "2-4 yrs", "LinkedIn/Public", "https://in.linkedin.com/in/manu-raj-sijaria", "", "SQL, Python, cloud data platforms, ETL/ELT, data warehousing.", "DM recruiter/contact if available", "", ""],
  ["Watch", 19, "Brewk Space", "Data Scientist", "Data Science", "Bengaluru", "On-site", "1-3 yrs", "LinkedIn/Public", "https://in.linkedin.com/company/brewkspaceindia", "hr@brewkspace.com", "Python, data analysis, statistical modeling, visualization, ML.", "Email resume if okay with on-site Bengaluru", "", ""],
];

const headers = [
  "Status", "Priority", "Company", "Role", "Track", "Location", "Mode", "Experience",
  "Source", "Apply Link", "Recruiter / Contact", "Why Match / Notes", "Next Action",
  "Applied Date", "Follow-up Date",
];

overview.showGridLines = false;
overview.getRange("A1:F1").merge();
overview.getRange("A1").values = [["Vraj Patel - AI/ML Job Search Tracker"]];
overview.getRange("A1").format = { font: { bold: true, color: "#FFFFFF", size: 18 }, fill: "#1F4E79" };
overview.getRange("A3:B8").values = [
  ["Total Jobs", jobs.length],
  ["To Review", jobs.filter(j => j[0] === "To Review").length],
  ["Watch", jobs.filter(j => j[0] === "Watch").length],
  ["Top Priority Cutoff", "1-10"],
  ["Preferred Locations", "Bangalore, Pune, Hyderabad, Mumbai, Remote/Hybrid"],
  ["Excluded Locations", "Ahmedabad, Vadodara, Gurgaon/Noida unless fully remote"],
];
overview.getRange("A3:A8").format = { fill: "#D9EAF7", font: { bold: true } };
overview.getRange("A10:F10").values = [["Suggested Application Order", "", "", "", "", ""]];
overview.getRange("A10:F10").merge();
overview.getRange("A10").format = { fill: "#E2F0D9", font: { bold: true } };
overview.getRange("A11:C20").values = jobs.slice(0, 10).map(j => [j[1], j[2], j[3]]);
overview.getRange("A11:C11").format = { font: { bold: true } };
overview.getRange("A3:F20").format.wrapText = true;
overview.getRange("A:A").format.columnWidthPx = 170;
overview.getRange("B:B").format.columnWidthPx = 250;
overview.getRange("C:C").format.columnWidthPx = 340;
overview.freezePanes.freezeRows(1);

tracker.showGridLines = false;
tracker.getRange("A1:O1").values = [headers];
tracker.getRange("A2:O20").values = jobs;
tracker.getRange("A1:O1").format = { fill: "#1F4E79", font: { bold: true, color: "#FFFFFF" }, wrapText: true };
tracker.getRange("A2:O20").format.wrapText = true;
tracker.getRange("B2:B20").format.numberFormat = "0";
tracker.getRange("N2:O20").setNumberFormat("yyyy-mm-dd");
tracker.tables.add("A1:O20", true, "JobsTracker");
tracker.freezePanes.freezeRows(1);
const widths = [110, 80, 180, 250, 140, 190, 110, 100, 130, 360, 200, 420, 280, 120, 130];
for (let i = 0; i < widths.length; i++) {
  tracker.getRangeByIndexes(0, i, 20, 1).format.columnWidthPx = widths[i];
}
tracker.getRange("A2:A20").dataValidation = { rule: { type: "list", values: ["To Review", "Applied", "Follow Up", "Interview", "Rejected", "Watch", "Paused"] } };
tracker.getRange("M2:M20").dataValidation = { rule: { type: "list", values: ["Open job and review Easy Apply form", "Review JD and tailor resume keywords", "Email resume and apply", "Apply on company careers page", "Apply via Foundit", "Verify active listing and apply", "DM recruiter/contact if available", "Follow up recruiter"] } };
tracker.getRange("A2:A20").conditionalFormats.add("containsText", { text: "To Review", format: { fill: "#FFF2CC" } });
tracker.getRange("A2:A20").conditionalFormats.add("containsText", { text: "Applied", format: { fill: "#D9EAD3" } });
tracker.getRange("A2:A20").conditionalFormats.add("containsText", { text: "Rejected", format: { fill: "#F4CCCC" } });
tracker.getRange("A2:A20").conditionalFormats.add("containsText", { text: "Watch", format: { fill: "#D9EAF7" } });

templates.showGridLines = false;
templates.getRange("A1:D1").merge();
templates.getRange("A1").values = [["Recruiter Outreach Templates"]];
templates.getRange("A1").format = { font: { bold: true, color: "#FFFFFF", size: 16 }, fill: "#1F4E79" };
templates.getRange("A3:B6").values = [
  ["LinkedIn Note", "Hi [Name],\n\nI am Vraj Patel, an AI/ML Software Engineer with 2.5 years of experience building production AI solutions including AWS Textract OCR pipelines, RAG chatbots, Neo4j knowledge graphs, Python microservices, data pipelines, and MLOps workflows.\n\nI saw the [Role] opening at [Company]. My experience with Python, LLM/RAG systems, LangChain/LangGraph, AWS, Neo4j, Docker, and production AI workflows looks closely aligned. I am available to join immediately and open to Bangalore, Pune, Hyderabad, Mumbai, or remote/hybrid roles.\n\nCould you please consider my profile for this role?\n\nRegards,\nVraj Patel"],
  ["Email Subject", "Application for [Role] - Vraj Patel - Immediate Joiner"],
  ["Email Body", "Hi [Name/Team],\n\nI am applying for the [Role] opening at [Company]. I have 2.5 years of experience in AI/ML and GenAI engineering, including OCR pipelines with AWS Textract, RAG applications, Neo4j knowledge graphs, Python microservices, data pipelines, and MLOps workflows.\n\nI am available to join immediately and open to Bangalore, Pune, Hyderabad, Mumbai, or remote/hybrid roles. Please find my resume attached.\n\nRegards,\nVraj Patel\npvraj1094@gmail.com"],
  ["Profile Keywords", "Python, SQL, LLMs, RAG, LangChain, LangGraph, Neo4j, AWS Textract, AWS, Azure ML, Docker, Kubernetes/OpenShift, MLOps, OCR, ETL, Data Engineering, NLP, PyTorch, TensorFlow"],
];
templates.getRange("A3:A6").format = { fill: "#D9EAF7", font: { bold: true }, wrapText: true };
templates.getRange("B3:B6").format.wrapText = true;
templates.getRange("A:A").format.columnWidthPx = 170;
templates.getRange("B:B").format.columnWidthPx = 760;

const inspect = await workbook.inspect({
  kind: "table",
  range: "Job Tracker!A1:O20",
  include: "values",
  tableMaxRows: 6,
  tableMaxCols: 8,
  maxChars: 3000,
});
console.log(inspect.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 50 },
  summary: "formula error scan",
  maxChars: 1000,
});
console.log(errors.ndjson);

await workbook.render({ sheetName: "Overview", autoCrop: "all", scale: 1, format: "png" });
await workbook.render({ sheetName: "Job Tracker", range: "A1:O12", scale: 1, format: "png" });
await workbook.render({ sheetName: "Templates", autoCrop: "all", scale: 1, format: "png" });

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(`${outputDir}/vraj_ai_ml_job_tracker.xlsx`);
console.log(`${outputDir}/vraj_ai_ml_job_tracker.xlsx`);
