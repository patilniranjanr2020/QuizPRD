/**
 * AI-Powered JRD (Joint Requirements Document) Creator
 * Tauri + Vite + TypeScript Single-Page Application
 */

import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType } from "docx";

// ============================================================================
// Types & State
// ============================================================================
interface GeminiModel {
  name: string;
  supportedGenerationMethods?: string[];
}

interface WizardState {
  apiKey: string;
  projectName: string;
  projectDescription: string;
  appType: "Web App" | "Desktop App" | "Android App" | "";
  attachedFileName: string;
  attachedFileContent: string;
  generatedJrd: string;
}

const state: WizardState = {
  apiKey: "",
  projectName: "",
  projectDescription: "",
  appType: "",
  attachedFileName: "",
  attachedFileContent: "",
  generatedJrd: ""
};

let debounceTimer: number | null = null;
let isValidatingKey = false;

// ============================================================================
// DOM Elements
// ============================================================================
// Step 1: API Key
const step1Card = document.getElementById("step-1-card") as HTMLElement;
const apiKeyInput = document.getElementById("api-key-input") as HTMLInputElement;
const inputContainer = document.getElementById("input-container") as HTMLElement;
const pasteActionBtn = document.getElementById("paste-action-btn") as HTMLButtonElement;
const toggleVisibilityBtn = document.getElementById("toggle-visibility-btn") as HTMLButtonElement;
const statusDot = document.getElementById("status-dot") as HTMLElement;
const statusText = document.getElementById("status-text") as HTMLElement;
const errorDetails = document.getElementById("error-details") as HTMLElement;
const errorCodeEl = document.getElementById("error-code") as HTMLElement;
const errorMessageText = document.getElementById("error-message-text") as HTMLElement;
const retryBtn = document.getElementById("retry-btn") as HTMLButtonElement;
const step1SummaryVal = document.getElementById("step-1-summary-val") as HTMLElement;
const editStep1Btn = document.getElementById("edit-step-1-btn") as HTMLButtonElement;
const askAiKeyBtn = document.getElementById("ask-ai-key-btn") as HTMLButtonElement;

// Animation Slot elements
const animIdle = document.getElementById("anim-idle") as HTMLElement;
const animLoading = document.getElementById("anim-loading") as HTMLElement;
const animSuccess = document.getElementById("anim-success") as HTMLElement;
const animError = document.getElementById("anim-error") as HTMLElement;

// Step 2: Project Name
const step2Card = document.getElementById("step-2-card") as HTMLElement;
const step2Form = document.getElementById("step-2-form") as HTMLFormElement;
const projectNameInput = document.getElementById("project-name-input") as HTMLInputElement;
const step2SummaryVal = document.getElementById("step-2-summary-val") as HTMLElement;
const editStep2Btn = document.getElementById("edit-step-2-btn") as HTMLButtonElement;
const askAiNameBtn = document.getElementById("ask-ai-name-btn") as HTMLButtonElement;

// Step 3: Project Description & File Attachment
const step3Card = document.getElementById("step-3-card") as HTMLElement;
const step3Form = document.getElementById("step-3-form") as HTMLFormElement;
const projectDescriptionInput = document.getElementById("project-description-input") as HTMLTextAreaElement;
const step3SummaryVal = document.getElementById("step-3-summary-val") as HTMLElement;
const editStep3Btn = document.getElementById("edit-step-3-btn") as HTMLButtonElement;
const askAiDescBtn = document.getElementById("ask-ai-desc-btn") as HTMLButtonElement;
const attachFileBtn = document.getElementById("attach-file-btn") as HTMLButtonElement;
const fileAttachmentInput = document.getElementById("file-attachment-input") as HTMLInputElement;
const attachedFilePill = document.getElementById("attached-file-pill") as HTMLElement;
const attachedFileName = document.getElementById("attached-file-name") as HTMLElement;
const removeFileBtn = document.getElementById("remove-file-btn") as HTMLButtonElement;

// Step 4: App Type
const step4Card = document.getElementById("step-4-card") as HTMLElement;
const appTypeRow = document.getElementById("app-type-row") as HTMLElement;
const generateJrdBtn = document.getElementById("generate-jrd-btn") as HTMLButtonElement;
const step4SummaryVal = document.getElementById("step-4-summary-val") as HTMLElement;
const editStep4Btn = document.getElementById("edit-step-4-btn") as HTMLButtonElement;

// Step 5: Result & Download
const step5ResultCard = document.getElementById("step-5-result-card") as HTMLElement;
const generatingView = document.getElementById("generating-view") as HTMLElement;
const documentView = document.getElementById("document-view") as HTMLElement;
const docTitleBadge = document.getElementById("doc-title-badge") as HTMLElement;
const docPlatformBadge = document.getElementById("doc-platform-badge") as HTMLElement;
const jrdContent = document.getElementById("jrd-content") as HTMLElement;
const downloadDocxBtn = document.getElementById("download-docx-btn") as HTMLButtonElement;
const copyJrdBtn = document.getElementById("copy-jrd-btn") as HTMLButtonElement;
const copyBtnLabel = document.getElementById("copy-btn-label") as HTMLElement;
const resetWizardBtn = document.getElementById("reset-wizard-btn") as HTMLButtonElement;
const workflowBadge = document.getElementById("workflow-badge") as HTMLElement;

// ============================================================================
// Animation & Status Controls
// ============================================================================
function setAnimationState(state: "idle" | "loading" | "success" | "error") {
  [animIdle, animLoading, animSuccess, animError].forEach((el) => {
    if (el) el.classList.remove("active");
  });

  inputContainer.classList.remove("state-loading", "state-valid", "state-invalid");

  switch (state) {
    case "idle":
      animIdle?.classList.add("active");
      break;
    case "loading":
      animLoading?.classList.add("active");
      inputContainer.classList.add("state-loading");
      break;
    case "success":
      animSuccess?.classList.add("active");
      inputContainer.classList.add("state-valid");
      break;
    case "error":
      animError?.classList.add("active");
      inputContainer.classList.add("state-invalid");
      break;
  }
}

function setStatus(text: string, type: "idle" | "loading" | "success" | "error") {
  if (statusText) statusText.textContent = text;
  if (statusDot) {
    statusDot.className = "status-indicator-dot";
    if (type !== "idle") statusDot.classList.add(type);
  }
}

function updateWorkflowBadge(stepText: string) {
  if (workflowBadge) workflowBadge.textContent = stepText;
}

function smoothScrollTo(element: HTMLElement) {
  setTimeout(() => {
    element.scrollIntoView({ behavior: "smooth", block: "center" });
  }, 100);
}

// ============================================================================
// STEP 1: API Key Verification Logic
// ============================================================================
async function validateApiKey(keyToTest: string): Promise<boolean> {
  const cleanKey = keyToTest.trim();

  if (!cleanKey) {
    setAnimationState("idle");
    setStatus("Waiting for API key... (starts validating automatically on paste)", "idle");
    errorDetails.classList.remove("visible");
    return false;
  }

  isValidatingKey = true;
  setAnimationState("loading");
  setStatus("Validating key with Google AI Studio in background...", "loading");
  errorDetails.classList.remove("visible");

  try {
    const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${encodeURIComponent(cleanKey)}`;
    const response = await fetch(url);

    if (response.ok) {
      state.apiKey = cleanKey;
      localStorage.setItem("gemini_api_key", cleanKey);

      setAnimationState("success");
      setStatus("Key verified & active! Moving to Step 2...", "success");

      // Complete Step 1 and automatically transition to Step 2
      setTimeout(() => {
        completeStep1(cleanKey);
      }, 700);

      isValidatingKey = false;
      return true;
    } else {
      let msg = "Google rejected this key. Please make sure the key has not expired.";
      try {
        const errorJson = await response.json();
        if (errorJson?.error?.message) {
          msg = errorJson.error.message;
        }
      } catch {
        // Fallback
      }

      setAnimationState("error");
      setStatus("Invalid API Key. Google rejected credentials.", "error");
      if (errorCodeEl) errorCodeEl.textContent = `HTTP ${response.status}`;
      if (errorMessageText) errorMessageText.textContent = msg;
      errorDetails.classList.add("visible");

      isValidatingKey = false;
      return false;
    }
  } catch (err: unknown) {
    setAnimationState("error");
    setStatus("Network connection error.", "error");
    if (errorCodeEl) errorCodeEl.textContent = "Network Error";
    if (errorMessageText) {
      errorMessageText.textContent = err instanceof Error ? err.message : "Failed to connect to Google API";
    }
    errorDetails.classList.add("visible");
    isValidatingKey = false;
    return false;
  }
}

function completeStep1(key: string) {
  step1Card.classList.add("completed");
  const maskedKey = key.length > 10 ? `${key.substring(0, 7)}...${key.substring(key.length - 4)}` : "Verified";
  if (step1SummaryVal) step1SummaryVal.textContent = maskedKey;

  // Reveal Step 2
  step2Card.classList.remove("hidden");
  step2Card.classList.add("anim-in", "active");
  updateWorkflowBadge("Step 2 of 4");
  smoothScrollTo(step2Card);
  setTimeout(() => projectNameInput?.focus(), 300);
}

// ============================================================================
// STEP 2: Project Name Logic
// ============================================================================
function completeStep2() {
  const name = projectNameInput.value.trim();
  if (!name) {
    projectNameInput.focus();
    projectNameInput.classList.add("state-invalid");
    setTimeout(() => projectNameInput.classList.remove("state-invalid"), 600);
    return;
  }

  state.projectName = name;
  step2Card.classList.remove("active");
  step2Card.classList.add("completed");
  if (step2SummaryVal) step2SummaryVal.textContent = name;

  // Reveal Step 3
  step3Card.classList.remove("hidden");
  step3Card.classList.add("anim-in", "active");
  updateWorkflowBadge("Step 3 of 4");
  smoothScrollTo(step3Card);
  setTimeout(() => projectDescriptionInput?.focus(), 300);
}

// ============================================================================
// STEP 3: Project Description Logic
// ============================================================================
function completeStep3() {
  const desc = projectDescriptionInput.value.trim();
  if (!desc) {
    projectDescriptionInput.focus();
    projectDescriptionInput.classList.add("state-invalid");
    setTimeout(() => projectDescriptionInput.classList.remove("state-invalid"), 600);
    return;
  }

  state.projectDescription = desc;
  step3Card.classList.remove("active");
  step3Card.classList.add("completed");
  const summarySnippet = desc.length > 70 ? `${desc.substring(0, 67)}...` : desc;
  if (step3SummaryVal) step3SummaryVal.textContent = summarySnippet;

  // Reveal Step 4
  step4Card.classList.remove("hidden");
  step4Card.classList.add("anim-in", "active");
  updateWorkflowBadge("Step 4 of 4");
  smoothScrollTo(step4Card);
}

// ============================================================================
// STEP 4: App Type Selection Logic (Single selection)
// ============================================================================
function selectAppType(type: "Web App" | "Desktop App" | "Android App") {
  state.appType = type;

  // Update UI cards - only one is selected
  document.querySelectorAll(".app-type-card").forEach((card) => {
    const cardType = card.getAttribute("data-type") as "Web App" | "Desktop App" | "Android App";
    const isSelected = cardType === type;
    card.classList.toggle("selected", isSelected);
    card.setAttribute("aria-checked", String(isSelected));
  });

  // Enable Generate Button
  generateJrdBtn.disabled = false;
}

// ============================================================================
// Gemini API Helper
// ============================================================================
async function fetchSupportedModels(key: string): Promise<string[]> {
  try {
    const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${encodeURIComponent(key)}`;
    const res = await fetch(url);
    if (!res.ok) return [];
    const data = await res.json();
    if (!Array.isArray(data.models)) return [];

    const valid = data.models
      .filter((m: GeminiModel) => m.supportedGenerationMethods?.includes("generateContent"))
      .map((m: GeminiModel) => m.name);

    valid.sort((a: string, b: string) => {
      const score = (name: string) => {
        const lower = name.toLowerCase();
        if (lower.includes("2.5-flash")) return 1;
        if (lower.includes("2.0-flash")) return 2;
        if (lower.includes("1.5-flash-latest")) return 3;
        if (lower.includes("1.5-flash")) return 4;
        if (lower.includes("1.5-pro-latest")) return 5;
        if (lower.includes("1.5-pro")) return 6;
        if (lower.includes("gemini")) return 7;
        return 10;
      };
      return score(a) - score(b);
    });

    return valid;
  } catch {
    return [];
  }
}

async function callGeminiQuick(promptText: string): Promise<string | null> {
  if (!state.apiKey) return null;
  const models = await fetchSupportedModels(state.apiKey);
  const modelsToTry = models.length > 0 ? models : [
    "models/gemini-2.0-flash",
    "models/gemini-1.5-flash",
    "models/gemini-1.5-flash-latest"
  ];

  for (const modelPath of modelsToTry) {
    const cleanPath = modelPath.startsWith("models/") ? modelPath : `models/${modelPath}`;
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/${cleanPath}:generateContent?key=${encodeURIComponent(state.apiKey)}`;
      const res = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          contents: [{ parts: [{ text: promptText }] }]
        })
      });
      const data = await res.json();
      if (res.ok && data?.candidates?.[0]?.content?.parts?.[0]?.text) {
        return data.candidates[0].content.parts[0].text.trim();
      }
    } catch {
      // Continue to next model
    }
  }
  return null;
}

// ============================================================================
// ASK AI: Smart Contextual Auto-Fill Feature
// ============================================================================

/**
 * Ask AI for API Key Field
 * Validates, guides, or checks clipboard
 */
async function handleAskAiKey() {
  const currentVal = apiKeyInput.value.trim();

  if (currentVal) {
    // If user already typed something, validate it directly
    validateApiKey(currentVal);
    return;
  }

  // If empty, check clipboard
  try {
    const clipboardText = await navigator.clipboard.readText();
    if (clipboardText && clipboardText.trim().startsWith("AIzaSy")) {
      apiKeyInput.value = clipboardText.trim();
      validateApiKey(clipboardText.trim());
      apiKeyInput.classList.add("ai-filled-highlight");
      setTimeout(() => apiKeyInput.classList.remove("ai-filled-highlight"), 1200);
      return;
    }
  } catch {
    // Clipboard permission not granted
  }

  // Show helpful guidance
  setStatus("To get an API key, visit Google AI Studio (aistudio.google.com/app/apikey). Free & takes 10 seconds!", "loading");
  apiKeyInput.focus();
}

/**
 * Ask AI for Project Name Field
 * Suggests or refines project name based on context
 */
async function handleAskAiName() {
  if (!state.apiKey) {
    alert("Please verify your Google AI Studio API key first in Step 1.");
    return;
  }

  const currentVal = projectNameInput.value.trim();
  const descVal = projectDescriptionInput.value.trim() || state.projectDescription;

  askAiNameBtn.classList.add("thinking");
  const originalText = askAiNameBtn.innerHTML;
  askAiNameBtn.innerHTML = `<span>Thinking...</span>`;

  let prompt = "";
  if (descVal && !currentVal) {
    prompt = `Suggest ONE catchy, modern, memorable product/project name for an application with this description: "${descVal}". Output ONLY the product name itself (2 to 4 words max), no quotes, no markdown, no explanation.`;
  } else if (currentVal) {
    prompt = `Refine and improve this project name to make it more professional, modern, and brandable: "${currentVal}". Context: "${descVal || "software application"}". Output ONLY the improved name itself, no quotes, no markdown.`;
  } else {
    prompt = `Suggest ONE creative, high-impact product name for a modern technology application (e.g. SwiftSync, PulseFlow, OmniTask). Output ONLY the product name itself, no quotes, no markdown.`;
  }

  const suggestedName = await callGeminiQuick(prompt);

  askAiNameBtn.classList.remove("thinking");
  askAiNameBtn.innerHTML = originalText;

  if (suggestedName) {
    // Clean up quotes or periods
    const cleanName = suggestedName.replace(/["'*.]/g, "").trim();
    projectNameInput.value = cleanName;
    projectNameInput.classList.add("ai-filled-highlight");
    setTimeout(() => projectNameInput.classList.remove("ai-filled-highlight"), 1200);
  } else {
    projectNameInput.value = "OmniFlow Pro";
  }
}

/**
 * Ask AI for Project Description Field
 * Generates or enriches project description based on context
 */
async function handleAskAiDesc() {
  if (!state.apiKey) {
    alert("Please verify your Google AI Studio API key first in Step 1.");
    return;
  }

  const projectName = projectNameInput.value.trim() || state.projectName || "Product";
  const currentDesc = projectDescriptionInput.value.trim();

  askAiDescBtn.classList.add("thinking");
  const originalText = askAiDescBtn.innerHTML;
  askAiDescBtn.innerHTML = `<span>Thinking...</span>`;

  const attachedContext = state.attachedFileContent
    ? `\nReference file context (${state.attachedFileName}):\n"${state.attachedFileContent.slice(0, 3000)}"\n`
    : "";

  let prompt = "";
  if (currentDesc) {
    prompt = `You are an expert product manager. Expand, improve, and polish this product description for "${projectName}": "${currentDesc}".${attachedContext} Write a comprehensive, well-structured 2-paragraph summary detailing the core purpose, target audience, primary user problem solved, and 3-4 key standout features. Write in natural professional language. Do not use markdown headers, write clean formatted paragraphs.`;
  } else if (projectName && projectName !== "Product") {
    prompt = `You are an expert product manager. Write a compelling, realistic product description for an application named "${projectName}".${attachedContext} Outline the core purpose, target audience, primary problem it solves, and 3-4 essential standout features. Write in natural, professional language in 2 paragraphs. Do not use markdown headers.`;
  } else {
    prompt = `You are an expert product manager. Write a compelling product description for a modern software application.${attachedContext} Outline the core purpose, target audience, primary problem it solves, and 3-4 key features in 2 natural paragraphs.`;
  }

  const generatedDesc = await callGeminiQuick(prompt);

  askAiDescBtn.classList.remove("thinking");
  askAiDescBtn.innerHTML = originalText;

  if (generatedDesc) {
    projectDescriptionInput.value = generatedDesc;
    projectDescriptionInput.classList.add("ai-filled-highlight");
    setTimeout(() => projectDescriptionInput.classList.remove("ai-filled-highlight"), 1200);
  }
}

// ============================================================================
// STEP 5: AI Generation Logic (Gemini API with JRD Structure)
// ============================================================================
async function generateJrdDocument() {
  if (!state.apiKey || !state.projectName || !state.projectDescription || !state.appType) {
    alert("Please complete all steps before generating the JRD document.");
    return;
  }

  // Complete Step 4 visually
  step4Card.classList.remove("active");
  step4Card.classList.add("completed");
  if (step4SummaryVal) step4SummaryVal.textContent = state.appType;

  // Reveal Step 5 Card & show Generating state
  step5ResultCard.classList.remove("hidden");
  step5ResultCard.classList.add("anim-in");
  generatingView.classList.remove("hidden");
  documentView.classList.add("hidden");
  updateWorkflowBadge("Generating JRD...");
  smoothScrollTo(step5ResultCard);

  // Discover working models
  let availableModels = await fetchSupportedModels(state.apiKey);
  if (availableModels.length === 0) {
    availableModels = [
      "models/gemini-2.0-flash",
      "models/gemini-1.5-flash",
      "models/gemini-1.5-flash-latest",
      "models/gemini-1.5-flash-8b",
      "models/gemini-1.5-pro",
      "models/gemini-1.5-pro-latest"
    ];
  }

  const attachedContext = state.attachedFileContent
    ? `\nATTACHED REFERENCE FILE CONTEXT (${state.attachedFileName}):\n"""\n${state.attachedFileContent.slice(0, 15000)}\n"""\n`
    : "";

  // Structured JRD Prompt as requested with all required sections:
  // Overview, Project Objective, Scope, Functional Requirements,
  // User Flows, Non-functional Requirements, Technical Considerations, Platform Selection
  const prompt = `You are a Principal Product Architect and Senior Technical Lead. Generate a comprehensive, professional Joint Requirements Document (JRD) based on the following verified specifications:

PROJECT NAME: ${state.projectName}
PROJECT DESCRIPTION: ${state.projectDescription}
TARGET APP PLATFORM: ${state.appType}
${attachedContext}

Please generate an exhaustive, professional, and well-structured JRD with the following mandatory sections:

# Joint Requirements Document (JRD): ${state.projectName}

## 1. Overview
- Executive Summary of the product vision, market problem, and value proposition.
- Target Application Platform: ${state.appType}.
- Document Version & Status: 1.0 (Draft).

## 2. Project Objective
- Core Problem Statement
- Strategic Business Objectives & Expected Outcomes
- Success Metrics & Key Performance Indicators (KPIs)

## 3. Scope
- In-Scope Features & Deliverables
- Out-of-Scope (Deferred to future releases)
- Target Stakeholders & Primary User Personas

## 4. Functional Requirements
- Epic 1: User Onboarding, Authentication & Profile Management
- Epic 2: Core Platform Functionality & Workflow Operations
- Epic 3: Data Management, Persistence, Sync & Reporting
- Epic 4: Administration, Configuration & Preferences
(Provide clear, actionable bulleted requirements with acceptance conditions for each epic)

## 5. User Flows
- Primary User Journey & Step-by-Step Flow
- Secondary/Exception Paths & Error Recovery
- Key Screen Transitions and Interaction Touchpoints

## 6. Non-functional Requirements
- Performance & SLA Latency Benchmarks (response SLAs, load times)
- Security & Encryption Standards (Authentication, Authorization, Data-at-rest & in-transit)
- Scalability & Resource Utilization
- Reliability & Disaster Recovery

## 7. Technical Considerations
- Platform Architecture Rationale (${state.appType})
- Recommended Client Framework, State Management & UI Library
- Backend, Database, and API Architecture
- Offline/Local Storage Strategy (if applicable)
- Third-party Integrations & External Services

## 8. Platform Selection
- Detailed Rationale for selecting ${state.appType}
- Platform-specific UX guidelines and device interaction requirements (e.g., responsive web design considerations, desktop windowing/offline caching, or Android material touch interactions)
- Implementation Roadmap (Phase 1 MVP, Phase 2 Beta, Phase 3 Production)
- Release Acceptance Criteria & Definition of Done

Make the document detailed, authoritative, structured with markdown headings, bullet points, and bold terms in natural, polished business language.`;

  let responseText = "";
  let usedModelName = "";

  for (const modelPath of availableModels) {
    const cleanPath = modelPath.startsWith("models/") ? modelPath : `models/${modelPath}`;
    usedModelName = cleanPath.replace("models/", "");

    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/${cleanPath}:generateContent?key=${encodeURIComponent(state.apiKey)}`;
      const res = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }]
        })
      });

      const data = await res.json();
      if (res.ok && data?.candidates?.[0]?.content?.parts?.[0]?.text) {
        responseText = data.candidates[0].content.parts[0].text;
        break;
      }
    } catch {
      // Continue to next model
    }
  }

  if (!responseText) {
    responseText = `# Joint Requirements Document (JRD): ${state.projectName}\n\n*Note: An error occurred while contacting the Gemini API.* Please verify your API key and network connection.`;
  }

  state.generatedJrd = responseText;

  // Render Result in Document View
  if (docTitleBadge) docTitleBadge.textContent = `${state.projectName} JRD`;
  if (docPlatformBadge) docPlatformBadge.textContent = `${state.appType} • ${usedModelName}`;
  renderJrdMarkdown(responseText);

  // Transition from generating to document view
  generatingView.classList.add("hidden");
  documentView.classList.remove("hidden");
  updateWorkflowBadge("JRD Generated ✓");
  smoothScrollTo(step5ResultCard);
}

// ============================================================================
// Markdown to HTML Renderer
// ============================================================================
function renderJrdMarkdown(markdown: string) {
  if (!jrdContent) return;

  const lines = markdown.split("\n");
  let html = "";
  let inList = false;

  for (const line of lines) {
    const trimmed = line.trim();

    if (!trimmed) {
      if (inList) {
        html += "</ul>";
        inList = false;
      }
      continue;
    }

    if (trimmed.startsWith("* ") || trimmed.startsWith("- ")) {
      if (!inList) {
        html += "<ul>";
        inList = true;
      }
      html += `<li>${formatInlineText(trimmed.substring(2))}</li>`;
    } else {
      if (inList) {
        html += "</ul>";
        inList = false;
      }

      if (trimmed.startsWith("# ")) {
        html += `<h1>${formatInlineText(trimmed.substring(2))}</h1>`;
      } else if (trimmed.startsWith("## ")) {
        html += `<h2>${formatInlineText(trimmed.substring(3))}</h2>`;
      } else if (trimmed.startsWith("### ")) {
        html += `<h3>${formatInlineText(trimmed.substring(4))}</h3>`;
      } else {
        html += `<p>${formatInlineText(trimmed)}</p>`;
      }
    }
  }

  if (inList) html += "</ul>";
  jrdContent.innerHTML = html;
}

function formatInlineText(str: string): string {
  let escaped = escapeHtml(str);
  escaped = escaped.replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>");
  escaped = escaped.replace(/`([^`]+)`/g, "<code>$1</code>");
  return escaped;
}

function escapeHtml(str: string): string {
  const div = document.createElement("div");
  div.textContent = str;
  return div.innerHTML;
}

// ============================================================================
// STEP 6: Word Document (.docx) Export using docx library
// ============================================================================
async function downloadWordDocument() {
  if (!state.generatedJrd) {
    alert("No JRD document generated yet.");
    return;
  }

  downloadDocxBtn.disabled = true;
  const originalText = downloadDocxBtn.innerHTML;
  downloadDocxBtn.innerHTML = `<span>Generating .docx...</span>`;

  try {
    const lines = state.generatedJrd.split("\n");
    const docChildren: Paragraph[] = [];

    // Title Block Header
    docChildren.push(
      new Paragraph({
        text: `Joint Requirements Document (JRD)`,
        heading: HeadingLevel.TITLE,
        alignment: AlignmentType.CENTER,
        spacing: { after: 120 }
      })
    );

    docChildren.push(
      new Paragraph({
        children: [
          new TextRun({ text: `Project: `, bold: true, size: 24 }),
          new TextRun({ text: `${state.projectName}  |  `, size: 24 }),
          new TextRun({ text: `Platform: `, bold: true, size: 24 }),
          new TextRun({ text: `${state.appType}  |  `, size: 24 }),
          new TextRun({ text: `Date: `, bold: true, size: 24 }),
          new TextRun({ text: new Date().toLocaleDateString(), size: 24 })
        ],
        alignment: AlignmentType.CENTER,
        spacing: { after: 360 }
      })
    );

    // Parse Markdown lines into docx Paragraphs
    for (const rawLine of lines) {
      const line = rawLine.trim();
      if (!line) continue;

      if (line.startsWith("# ")) {
        docChildren.push(
          new Paragraph({
            text: line.substring(2),
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 280, after: 120 }
          })
        );
      } else if (line.startsWith("## ")) {
        docChildren.push(
          new Paragraph({
            text: line.substring(3),
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 240, after: 100 }
          })
        );
      } else if (line.startsWith("### ")) {
        docChildren.push(
          new Paragraph({
            text: line.substring(4),
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 80 }
          })
        );
      } else if (line.startsWith("- ") || line.startsWith("* ")) {
        // Bullet Point with bold inline parsing
        const itemContent = line.substring(2);
        docChildren.push(
          new Paragraph({
            children: parseInlineFormattingToRuns(itemContent),
            bullet: { level: 0 },
            spacing: { before: 60, after: 60 }
          })
        );
      } else {
        // Regular paragraph with bold inline parsing
        docChildren.push(
          new Paragraph({
            children: parseInlineFormattingToRuns(line),
            spacing: { before: 80, after: 100 }
          })
        );
      }
    }

    // Create Microsoft Word Document instance
    const doc = new Document({
      sections: [
        {
          properties: {},
          children: docChildren
        }
      ]
    });

    // Pack into Blob and trigger browser download
    const blob = await Packer.toBlob(doc);
    const safeFilename = `${state.projectName.replace(/[^a-zA-Z0-9_-]/g, "_")}_JRD.docx`;

    const downloadUrl = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = downloadUrl;
    link.download = safeFilename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(downloadUrl);

  } catch (err: unknown) {
    console.error("Error generating docx:", err);
    alert(`Failed to generate Word document: ${err instanceof Error ? err.message : String(err)}`);
  } finally {
    downloadDocxBtn.disabled = false;
    downloadDocxBtn.innerHTML = originalText;
  }
}

function parseInlineFormattingToRuns(text: string): TextRun[] {
  const runs: TextRun[] = [];
  const parts = text.split(/(\*\*.*?\*\*)/g);

  for (const part of parts) {
    if (!part) continue;
    if (part.startsWith("**") && part.endsWith("**")) {
      runs.push(
        new TextRun({
          text: part.slice(2, -2),
          bold: true
        })
      );
    } else {
      runs.push(
        new TextRun({
          text: part
        })
      );
    }
  }

  return runs;
}

// ============================================================================
// Reset / Edit Controls
// ============================================================================
function resetWizard() {
  state.projectName = "";
  state.projectDescription = "";
  state.appType = "";
  state.attachedFileName = "";
  state.attachedFileContent = "";
  state.generatedJrd = "";

  projectNameInput.value = "";
  projectDescriptionInput.value = "";
  if (fileAttachmentInput) fileAttachmentInput.value = "";
  attachedFilePill?.classList.add("hidden");
  document.querySelectorAll(".app-type-card").forEach((c) => c.classList.remove("selected"));
  generateJrdBtn.disabled = true;

  step2Card.classList.remove("completed");
  step3Card.classList.remove("completed");
  step4Card.classList.remove("completed");

  step2Card.classList.add("hidden");
  step3Card.classList.add("hidden");
  step4Card.classList.add("hidden");
  step5ResultCard.classList.add("hidden");

  // Keep API Key verified, advance to Step 2
  if (state.apiKey) {
    step2Card.classList.remove("hidden");
    step2Card.classList.add("anim-in", "active");
    updateWorkflowBadge("Step 2 of 4");
    smoothScrollTo(step2Card);
    projectNameInput.focus();
  } else {
    step1Card.classList.remove("completed");
    updateWorkflowBadge("Step 1 of 4");
    smoothScrollTo(step1Card);
    apiKeyInput.focus();
  }
}

// ============================================================================
// Lifecycle & Event Listeners
// ============================================================================
window.addEventListener("DOMContentLoaded", () => {
  // Check for stored API key
  const savedKey = localStorage.getItem("gemini_api_key");
  if (savedKey) {
    apiKeyInput.value = savedKey;
    validateApiKey(savedKey);
  } else {
    apiKeyInput.focus();
  }

  // API Key Typing Listener with Debounce
  apiKeyInput?.addEventListener("input", (e) => {
    const val = (e.target as HTMLInputElement).value;
    if (debounceTimer) clearTimeout(debounceTimer);

    if (!val.trim()) {
      setAnimationState("idle");
      setStatus("Waiting for API key... (starts validating automatically on paste)", "idle");
      errorDetails.classList.remove("visible");
      return;
    }

    setAnimationState("loading");
    setStatus("Analyzing key...", "loading");

    debounceTimer = window.setTimeout(() => {
      validateApiKey(val);
    }, 450);
  });

  // Instant Paste Listener
  apiKeyInput?.addEventListener("paste", () => {
    setTimeout(() => {
      if (apiKeyInput.value.trim()) {
        validateApiKey(apiKeyInput.value.trim());
      }
    }, 20);
  });

  // Enter Key on API Key Input
  apiKeyInput?.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      if (!isValidatingKey) {
        validateApiKey(apiKeyInput.value);
      }
    }
  });

  // Clipboard Paste Button
  pasteActionBtn?.addEventListener("click", async () => {
    try {
      const text = await navigator.clipboard.readText();
      if (text) {
        apiKeyInput.value = text.trim();
        validateApiKey(text.trim());
      }
    } catch {
      apiKeyInput.focus();
    }
  });

  // Toggle Password Visibility
  toggleVisibilityBtn?.addEventListener("click", () => {
    if (apiKeyInput.type === "password") {
      apiKeyInput.type = "text";
      toggleVisibilityBtn.setAttribute("title", "Hide Key");
    } else {
      apiKeyInput.type = "password";
      toggleVisibilityBtn.setAttribute("title", "Show Key");
    }
  });

  // Retry API Key Check
  retryBtn?.addEventListener("click", () => {
    validateApiKey(apiKeyInput.value);
  });

  // Edit Step 1 (Change API Key)
  editStep1Btn?.addEventListener("click", () => {
    step1Card.classList.remove("completed");
    step1Card.classList.add("active");
    apiKeyInput.focus();
    smoothScrollTo(step1Card);
  });

  // Ask AI Button for API Key
  askAiKeyBtn?.addEventListener("click", () => {
    handleAskAiKey();
  });

  // Step 2: Project Name Submission
  step2Form?.addEventListener("submit", (e) => {
    e.preventDefault();
    completeStep2();
  });

  // Edit Step 2
  editStep2Btn?.addEventListener("click", () => {
    step2Card.classList.remove("completed");
    step2Card.classList.add("active");
    projectNameInput.focus();
    smoothScrollTo(step2Card);
  });

  // Ask AI Button for Project Name
  askAiNameBtn?.addEventListener("click", () => {
    handleAskAiName();
  });

  // Step 3: Project Description Submission
  step3Form?.addEventListener("submit", (e) => {
    e.preventDefault();
    completeStep3();
  });

  // Ctrl+Enter inside Description Textarea submits
  projectDescriptionInput?.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
      e.preventDefault();
      completeStep3();
    }
  });

  // Quick Inspiration Chips in Step 3
  document.querySelectorAll(".sample-chip").forEach((chip) => {
    chip.addEventListener("click", () => {
      const sampleText = chip.getAttribute("data-text");
      if (sampleText) {
        projectDescriptionInput.value = sampleText;
        projectDescriptionInput.focus();
      }
    });
  });

  // Edit Step 3
  editStep3Btn?.addEventListener("click", () => {
    step3Card.classList.remove("completed");
    step3Card.classList.add("active");
    projectDescriptionInput.focus();
    smoothScrollTo(step3Card);
  });

  // Ask AI Button for Project Description
  askAiDescBtn?.addEventListener("click", () => {
    handleAskAiDesc();
  });

  // File Attachment Controls
  attachFileBtn?.addEventListener("click", () => {
    fileAttachmentInput?.click();
  });

  fileAttachmentInput?.addEventListener("change", (e) => {
    const target = e.target as HTMLInputElement;
    const file = target.files?.[0];
    if (!file) return;

    state.attachedFileName = file.name;
    if (attachedFileName) attachedFileName.textContent = file.name;
    attachedFilePill?.classList.remove("hidden");

    const reader = new FileReader();
    reader.onload = (event) => {
      const content = event.target?.result;
      if (typeof content === "string") {
        state.attachedFileContent = content;
      } else {
        state.attachedFileContent = `[Attached reference document: ${file.name}, size: ${file.size} bytes]`;
      }
    };
    reader.onerror = () => {
      state.attachedFileContent = `[Attached reference file: ${file.name}]`;
    };
    reader.readAsText(file);
  });

  removeFileBtn?.addEventListener("click", () => {
    state.attachedFileName = "";
    state.attachedFileContent = "";
    if (fileAttachmentInput) fileAttachmentInput.value = "";
    attachedFilePill?.classList.add("hidden");
  });

  // Step 4: App Type Selection (Single selection)
  appTypeRow?.addEventListener("click", (e) => {
    const target = e.target as HTMLElement;
    const card = target.closest(".app-type-card") as HTMLElement | null;
    if (card) {
      const type = card.getAttribute("data-type") as "Web App" | "Desktop App" | "Android App";
      if (type) {
        selectAppType(type);
      }
    }
  });

  // Edit Step 4
  editStep4Btn?.addEventListener("click", () => {
    step4Card.classList.remove("completed");
    step4Card.classList.add("active");
    smoothScrollTo(step4Card);
  });

  // Step 5: Generate JRD Button
  generateJrdBtn?.addEventListener("click", () => {
    generateJrdDocument();
  });

  // Step 6: Download Word Document (.docx)
  downloadDocxBtn?.addEventListener("click", () => {
    downloadWordDocument();
  });

  // Copy JRD Text to Clipboard
  copyJrdBtn?.addEventListener("click", () => {
    if (state.generatedJrd) {
      navigator.clipboard.writeText(state.generatedJrd);
      if (copyBtnLabel) copyBtnLabel.textContent = "Copied!";
      setTimeout(() => {
        if (copyBtnLabel) copyBtnLabel.textContent = "Copy Text";
      }, 1800);
    }
  });

  // Reset Wizard
  resetWizardBtn?.addEventListener("click", resetWizard);
});
