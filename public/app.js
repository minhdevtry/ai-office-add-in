/**
 * AI Word Taskpane - Main Application Controller (2026 Premium Edition)
 */

/* global Office */

import { wordBridge } from "./modules/word-bridge.js";
import { docState } from "./modules/doc-state.js";
import { sseClient } from "./modules/sse-client.js";
import { wsClient } from "./modules/ws-client.js";

class AppController {
  constructor() {
    this.state = {
      messages: [],
      activeContext: null, // { type: 'selection' | 'document', text: string, words: number }
      isStreaming: false,
      skills: [],
      theme: "light",
      config: {
        endpoint: "",
        apiKey: "",
        model: "",
        enableThinking: true,
        houseVoice: "",
      },
    };

    this.dom = {};
  }

  async init() {
    this.cacheDom();
    this.initTheme();
    this.bindEvents();
    await this.loadSkills();
    this.initOffice();
    this.initWebSocket();
    this.loadState();
    this.renderSkills();
  }

  cacheDom() {
    this.dom = {
      statusPill: document.getElementById("statusPill"),
      statusText: document.getElementById("statusText"),
      btnToggleTheme: document.getElementById("btnToggleTheme"),
      btnSettings: document.getElementById("btnSettings"),
      skillsBar: document.getElementById("skillsBar"),
      contextBar: document.getElementById("contextBar"),
      contextTypeIcon: document.getElementById("contextTypeIcon"),
      contextLabel: document.getElementById("contextLabel"),
      btnClearContext: document.getElementById("btnClearContext"),
      agentBanner: document.getElementById("agentBanner"),
      agentBannerText: document.getElementById("agentBannerText"),
      chatContainer: document.getElementById("chatContainer"),
      emptyHint: document.getElementById("emptyHint"),
      btnReadSelection: document.getElementById("btnReadSelection"),
      btnReadDoc: document.getElementById("btnReadDoc"),
      btnClearChat: document.getElementById("btnClearChat"),
      chatInput: document.getElementById("chatInput"),
      btnSend: document.getElementById("btnSend"),
      // Settings Modal
      settingsModal: document.getElementById("settingsModal"),
      btnCloseSettings: document.getElementById("btnCloseSettings"),
      btnSaveSettings: document.getElementById("btnSaveSettings"),
      btnTestConn: document.getElementById("btnTestConn"),
      cfgEndpoint: document.getElementById("cfgEndpoint"),
      cfgApiKey: document.getElementById("cfgApiKey"),
      cfgModel: document.getElementById("cfgModel"),
      cfgHouseVoice: document.getElementById("cfgHouseVoice"),
    };
  }

  initTheme() {
    const savedTheme = localStorage.getItem("ai_word.theme") || "light";
    this.setTheme(savedTheme);
  }

  setTheme(theme) {
    this.state.theme = theme;
    document.documentElement.setAttribute("data-theme", theme);
    localStorage.setItem("ai_word.theme", theme);
    if (this.dom.btnToggleTheme) {
      this.dom.btnToggleTheme.textContent = theme === "dark" ? "☀️" : "🌙";
      this.dom.btnToggleTheme.title = theme === "dark" ? "Chuyển sang Giao diện Sáng" : "Chuyển sang Giao diện Tối";
    }
  }

  toggleTheme() {
    const nextTheme = this.state.theme === "dark" ? "light" : "dark";
    this.setTheme(nextTheme);
  }

  bindEvents() {
    // Theme toggle
    this.dom.btnToggleTheme.addEventListener("click", () => this.toggleTheme());

    // Composer events
    this.dom.btnSend.addEventListener("click", () => this.handleSendOrStop());
    this.dom.chatInput.addEventListener("keydown", (e) => {
      if (e.key === "Enter" && !e.shiftKey) {
        e.preventDefault();
        this.handleSendOrStop();
      }
    });

    // Auto-expand textarea
    this.dom.chatInput.addEventListener("input", () => {
      this.dom.chatInput.style.height = "auto";
      this.dom.chatInput.style.height = Math.min(this.dom.chatInput.scrollHeight, 120) + "px";
    });

    // Toolbar buttons
    this.dom.btnReadSelection.addEventListener("click", () => this.captureSelection());
    this.dom.btnReadDoc.addEventListener("click", () => this.captureFullDocument());
    this.dom.btnClearContext.addEventListener("click", () => this.clearContext());
    this.dom.btnClearChat.addEventListener("click", () => this.clearChatHistory());

    // Starter Cards click delegation
    this.dom.emptyHint.addEventListener("click", (e) => {
      const card = e.target.closest(".starter-card");
      if (card) {
        const skillId = card.dataset.skill;
        const skill = this.state.skills.find((s) => s.id === skillId);
        if (skill) this.triggerSkill(skill);
      }
    });

    // Settings Modal
    this.dom.btnSettings.addEventListener("click", () => this.openSettings());
    this.dom.btnCloseSettings.addEventListener("click", () => this.closeSettings());
    this.dom.btnSaveSettings.addEventListener("click", () => this.saveSettings());
    this.dom.btnTestConn.addEventListener("click", () => this.testAiConnection());
    this.dom.settingsModal.addEventListener("click", (e) => {
      if (e.target === this.dom.settingsModal) this.closeSettings();
    });
  }

  async loadSkills() {
    try {
      const res = await fetch("skills.json");
      this.state.skills = await res.json();
    } catch (_) {
      this.state.skills = [];
    }
  }

  renderSkills() {
    this.dom.skillsBar.innerHTML = "";
    for (const skill of this.state.skills) {
      const pill = document.createElement("button");
      pill.className = "skill-pill";
      pill.title = skill.tip;
      pill.innerHTML = `<span>${skill.icon}</span> <span>${skill.label}</span>`;
      pill.addEventListener("click", () => this.triggerSkill(skill));
      this.dom.skillsBar.appendChild(pill);
    }
  }

  initOffice() {
    if (typeof Office !== "undefined") {
      Office.onReady((info) => {
        if (info.host === Office.HostType.Word) {
          this.updateStatus("online", "Word Ready");
        } else {
          this.updateStatus("busy", "Web Preview");
        }
      });
    } else {
      this.updateStatus("busy", "Standalone");
    }
  }

  initWebSocket() {
    wsClient.onStatusChange = (status) => {
      if (status === "connected") {
        this.updateStatus("online", "Bridge OK");
      } else if (status === "connecting") {
        this.updateStatus("busy", "Nối lại...");
      } else {
        this.updateStatus("offline", "Mất nối");
      }
    };

    wsClient.onAgentActive = (active, toolName) => {
      if (active) {
        this.dom.agentBanner.classList.add("show");
        this.dom.agentBannerText.textContent = `AI Agent đang thực thi: ${toolName}...`;
      } else {
        this.dom.agentBanner.classList.remove("show");
      }
    };

    wsClient.connect();
  }

  updateStatus(stateClass, text) {
    this.dom.statusPill.className = `status-pill ${stateClass}`;
    this.dom.statusText.textContent = text;
  }

  loadState() {
    this.state.messages = docState.loadHistory();
    this.state.config.houseVoice = docState.loadHouseVoice();
    const cfg = docState.loadConfig();
    this.state.config = { ...this.state.config, ...cfg };

    if (this.state.messages.length > 0) {
      this.dom.emptyHint.style.display = "none";
      for (const msg of this.state.messages) {
        this.renderMessageBubble(msg.role, msg.content, {
          thinking: msg.thinking,
          persist: false,
        });
      }
    }
  }

  // ─── CONTEXT CAPTURE ───

  async captureSelection() {
    try {
      const text = await wordBridge.getSelectedText();
      if (!text || text.trim() === "") {
        this.showToast("⚠️ Chưa có đoạn văn nào được bôi đen trong Word.");
        return;
      }
      const analytics = wordBridge.getTextAnalytics(text);
      this.setContext("selection", text, `Bôi đen: ${analytics.words} từ`);
      this.showToast(`🎯 Đã nạp vùng chọn (${analytics.words} từ)`);
    } catch (err) {
      this.showToast(`⚠️ Lỗi đọc vùng chọn: ${err.message}`);
    }
  }

  async captureFullDocument() {
    try {
      const text = await wordBridge.getFullDocumentText(20000);
      if (!text || text.trim() === "") {
        this.showToast("⚠️ Tài liệu hiện đang trống.");
        return;
      }
      const analytics = wordBridge.getTextAnalytics(text);
      this.setContext("document", text, `Toàn bài: ${analytics.words} từ`);
      this.showToast(`📑 Đã nạp toàn bộ bài (${analytics.words} từ)`);
    } catch (err) {
      this.showToast(`⚠️ Lỗi đọc tài liệu: ${err.message}`);
    }
  }

  setContext(type, text, label) {
    this.state.activeContext = { type, text, words: text.split(/\s+/).length };
    this.dom.contextBar.style.display = "flex";
    this.dom.contextTypeIcon.textContent = type === "selection" ? "🎯" : "📑";
    this.dom.contextLabel.textContent = label;
  }

  clearContext() {
    this.state.activeContext = null;
    this.dom.contextBar.style.display = "none";
  }

  // ─── SKILLS TRIGGER ───

  async triggerSkill(skill) {
    // Tự động nạp vùng chọn nếu chưa có
    if (!this.state.activeContext) {
      const selected = await wordBridge.getSelectedText();
      if (selected && selected.trim()) {
        this.setContext("selection", selected, `Bôi đen: ${selected.split(/\s+/).length} từ`);
      }
    }

    let promptText = skill.prompt;
    if (this.state.activeContext) {
      promptText = `${skill.prompt}\n\n[Đoạn văn bản mục tiêu]:\n"""\n${this.state.activeContext.text}\n"""`;
    }

    this.sendMessage(promptText, { skillId: skill.id, defaultMode: skill.defaultMode });
  }

  // ─── CHAT & STREAMING ───

  handleSendOrStop() {
    if (this.state.isStreaming) {
      sseClient.abort();
      this.setStreamingState(false);
      return;
    }

    const text = this.dom.chatInput.value.trim();
    if (!text && !this.state.activeContext) return;

    let fullPrompt = text;
    if (this.state.activeContext && !text.includes(this.state.activeContext.text)) {
      fullPrompt = `${text ? text + "\n\n" : ""}[Ngữ cảnh văn bản đính kèm]:\n"""\n${this.state.activeContext.text}\n"""`;
    }

    this.dom.chatInput.value = "";
    this.dom.chatInput.style.height = "auto";
    this.sendMessage(fullPrompt);
  }

  async sendMessage(promptText, meta = {}) {
    this.dom.emptyHint.style.display = "none";

    // 1. Render User Message
    this.renderMessageBubble("user", promptText);
    this.state.messages.push({ role: "user", content: promptText });
    docState.saveHistory(this.state.messages);

    // 2. Chuẩn bị Assistant Bubble
    const assistantBubble = this.createAssistantBubble();
    this.setStreamingState(true);

    // 3. Chuẩn bị System Prompt & House Voice
    let systemPrompt = docState.loadSystemPrompt();
    const houseVoice = this.state.config.houseVoice || docState.loadHouseVoice();
    if (houseVoice && houseVoice.trim()) {
      systemPrompt += `\n\n<personal_writing_style>\n${houseVoice.trim()}\nQuy tắc: Luôn mô phỏng đúng văn phong, nhịp điệu và cách dùng từ của đoạn mẫu trên.\n</personal_writing_style>`;
    }

    try {
      await sseClient.streamChat({
        messages: this.state.messages,
        systemPrompt,
        endpoint: this.state.config.endpoint,
        apiKey: this.state.config.apiKey,
        model: this.state.config.model,
        enableThinking: this.state.config.enableThinking !== false,

        onThinking: (chunk, fullThinking) => {
          assistantBubble.updateThinking(fullThinking);
          this.scrollToBottom();
        },

        onDelta: (chunk, fullText) => {
          assistantBubble.updateContent(fullText, true);
          this.scrollToBottom();
        },

        onDone: (fullText, fullThinking, aborted) => {
          this.setStreamingState(false);
          assistantBubble.finalize(fullText, fullThinking, meta);

          this.state.messages.push({
            role: "assistant",
            content: fullText,
            thinking: fullThinking,
          });
          docState.saveHistory(this.state.messages);
          this.scrollToBottom();
        },

        onError: (err) => {
          this.setStreamingState(false);
          assistantBubble.showError(err.message);
          this.scrollToBottom();
        },
      });
    } catch (_) {
      this.setStreamingState(false);
    }
  }

  setStreamingState(isStreaming) {
    this.state.isStreaming = isStreaming;
    if (isStreaming) {
      this.dom.btnSend.classList.add("stop");
      this.dom.btnSend.innerHTML = "<span>■</span>";
      this.dom.btnSend.title = "Dừng sinh";
    } else {
      this.dom.btnSend.classList.remove("stop");
      this.dom.btnSend.innerHTML = "<span>➤</span>";
      this.dom.btnSend.title = "Gửi yêu cầu";
    }
  }

  // ─── RENDERING BUBBLES ───

  renderMessageBubble(role, content, { thinking = "", persist = true } = {}) {
    const bubble = document.createElement("div");
    bubble.className = `message-bubble ${role}`;

    if (role === "user") {
      bubble.textContent = content;
    } else if (role === "assistant") {
      if (thinking) {
        const thinkingBox = document.createElement("details");
        thinkingBox.className = "thinking-box";
        thinkingBox.innerHTML = `<summary>🧠 Quá trình suy luận</summary><pre class="thinking-content">${this.escapeHtml(thinking)}</pre>`;
        bubble.appendChild(thinkingBox);
      }
      const textDiv = document.createElement("div");
      textDiv.className = "assistant-content";
      textDiv.innerHTML = this.formatMarkdown(content);
      bubble.appendChild(textDiv);

      // Thêm action buttons
      this.attachActionButtons(bubble, content);
      this.attachAnchorChips(bubble, content);
    }

    this.dom.chatContainer.appendChild(bubble);
    this.scrollToBottom();
    return bubble;
  }

  createAssistantBubble() {
    const bubble = document.createElement("div");
    bubble.className = "message-bubble assistant";

    let thinkingBox = null;
    let thinkingContent = null;
    const contentDiv = document.createElement("div");
    contentDiv.className = "assistant-content";
    bubble.appendChild(contentDiv);

    this.dom.chatContainer.appendChild(bubble);

    return {
      updateThinking: (text) => {
        if (!thinkingBox) {
          thinkingBox = document.createElement("details");
          thinkingBox.className = "thinking-box";
          thinkingBox.open = true;
          thinkingBox.innerHTML = `<summary>🧠 Đang suy luận...</summary><pre class="thinking-content"></pre>`;
          bubble.insertBefore(thinkingBox, contentDiv);
          thinkingContent = thinkingBox.querySelector(".thinking-content");
        }
        if (thinkingContent) thinkingContent.textContent = text;
      },

      updateContent: (text, isStreaming = false) => {
        if (thinkingBox && thinkingBox.open) {
          thinkingBox.open = false;
          const sum = thinkingBox.querySelector("summary");
          if (sum) sum.textContent = "🧠 Quá trình suy luận (Nhấp để mở)";
        }
        contentDiv.innerHTML = this.formatMarkdown(text) + (isStreaming ? '<span class="streaming-cursor"></span>' : '');
      },

      finalize: (fullText, fullThinking, meta) => {
        contentDiv.innerHTML = this.formatMarkdown(fullText);
        this.attachActionButtons(bubble, fullText, meta);
        this.attachAnchorChips(bubble, fullText);
      },

      showError: (errText) => {
        contentDiv.innerHTML += `<div style="color: var(--danger); margin-top: 6px; font-weight: 500;">⚠️ ${this.escapeHtml(errText)}</div>`;
      },
    };
  }

  // ─── ACTION BUTTONS (4 INSERT MODES + COMMENT + COPY) ───

  attachActionButtons(container, text, meta = {}) {
    const actionsBar = document.createElement("div");
    actionsBar.className = "msg-actions";

    const actions = [
      { id: "replace", label: "🔄 Thay thế", primary: true, fn: () => wordBridge.replaceSelection(text) },
      { id: "cursor", label: "➕ Tại con trỏ", fn: () => wordBridge.insertAtCursor(text) },
      { id: "after", label: "⬇️ Sau đoạn", fn: () => wordBridge.insertParagraphAfter(text) },
      { id: "end", label: "📄 Cuối trang", fn: () => wordBridge.appendDocumentEnd(text) },
      { id: "comment", label: "💬 Bình luận lề", fn: () => wordBridge.insertComment(text) },
      {
        id: "copy",
        label: "📋 Sao chép",
        fn: () => {
          navigator.clipboard.writeText(text);
          this.showToast("📋 Đã sao chép vào bộ nhớ tạm!");
        },
      },
    ];

    for (const act of actions) {
      const btn = document.createElement("button");
      btn.className = `action-btn ${act.primary ? "primary" : ""}`;
      btn.textContent = act.label;
      btn.addEventListener("click", async () => {
        try {
          await act.fn();
          this.showToast(`✅ Đã thực hiện: ${act.label}`);
        } catch (err) {
          this.showToast(`⚠️ Lỗi: ${err.message}`);
        }
      });
      actionsBar.appendChild(btn);
    }

    container.appendChild(actionsBar);
  }

  // ─── ANCHOR CHIPS (CLICK-TO-JUMP) ───

  attachAnchorChips(container, text) {
    // Tìm các câu trích dẫn trong ngoặc kép hoặc blockquote
    const quoteMatches = text.match(/"([^"]{15,100})"/g) || [];

    if (quoteMatches.length > 0) {
      const chipsContainer = document.createElement("div");
      chipsContainer.style.marginTop = "6px";

      for (const quote of quoteMatches.slice(0, 3)) {
        const clean = quote.replace(/^"|"$/g, "").trim();
        const chip = document.createElement("button");
        chip.className = "anchor-chip";
        chip.innerHTML = `<span>🎯 Xem trong Word:</span> <em>"${clean.slice(0, 32)}..."</em>`;
        chip.addEventListener("click", async () => {
          try {
            await wordBridge.selectRange(clean);
            chip.classList.add("jump-success");
            this.showToast("🎯 Đã định vị và bôi đen đoạn văn trong Word!");
            setTimeout(() => chip.classList.remove("jump-success"), 1500);
          } catch (err) {
            this.showToast(`⚠️ ${err.message}`);
          }
        });
        chipsContainer.appendChild(chip);
      }
      container.insertBefore(chipsContainer, container.querySelector(".msg-actions"));
    }
  }

  // ─── SETTINGS MODAL & TEST CONNECTION ───

  openSettings() {
    this.dom.cfgEndpoint.value = this.state.config.endpoint || "";
    this.dom.cfgApiKey.value = this.state.config.apiKey || "";
    this.dom.cfgModel.value = this.state.config.model || "";
    this.dom.cfgHouseVoice.value = this.state.config.houseVoice || "";
    this.dom.settingsModal.classList.add("open");
  }

  closeSettings() {
    this.dom.settingsModal.classList.remove("open");
  }

  saveSettings() {
    this.state.config.endpoint = this.dom.cfgEndpoint.value.trim();
    this.state.config.apiKey = this.dom.cfgApiKey.value.trim();
    this.state.config.model = this.dom.cfgModel.value.trim();
    this.state.config.houseVoice = this.dom.cfgHouseVoice.value.trim();

    docState.saveConfig(this.state.config);
    docState.saveHouseVoice(this.state.config.houseVoice);

    this.closeSettings();
    this.showToast("✅ Đã lưu cài đặt thành công!");
  }

  async testAiConnection() {
    this.dom.btnTestConn.textContent = "Đang kiểm tra...";
    this.dom.btnTestConn.disabled = true;

    const endpoint = this.dom.cfgEndpoint.value.trim();
    const apiKey = this.dom.cfgApiKey.value.trim();
    const model = this.dom.cfgModel.value.trim();

    try {
      const res = await fetch("/api/chat", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          messages: [{ role: "user", content: "Chào bạn, hãy trả lời 'OK'." }],
          endpoint,
          apiKey,
          model,
          stream: false,
          maxTokens: 10,
        }),
      });

      if (res.ok) {
        this.showToast("✨ Kết nối AI Endpoint thành công!");
      } else {
        const errJson = await res.json().catch(() => ({}));
        this.showToast(`⚠️ Lỗi kết nối: ${errJson.error || res.statusText}`);
      }
    } catch (err) {
      this.showToast(`⚠️ Không thể kết nối: ${err.message}`);
    } finally {
      this.dom.btnTestConn.textContent = "Kiểm tra kết nối";
      this.dom.btnTestConn.disabled = false;
    }
  }

  clearChatHistory() {
    if (confirm("Bạn có chắc chắn muốn xóa toàn bộ lịch sử chat của tài liệu này?")) {
      this.state.messages = [];
      docState.clearHistory();
      this.dom.chatContainer.innerHTML = "";
      this.dom.chatContainer.appendChild(this.dom.emptyHint);
      this.dom.emptyHint.style.display = "flex";
      this.showToast("🗑️ Đã xóa lịch sử trò chuyện.");
    }
  }

  // ─── RICH MARKDOWN PARSER ───

  formatMarkdown(text) {
    if (!text) return "";
    let html = this.escapeHtml(text);

    // Headers
    html = html.replace(/^### (.*$)/gim, "<h3>$1</h3>");
    html = html.replace(/^## (.*$)/gim, "<h2>$1</h2>");
    html = html.replace(/^# (.*$)/gim, "<h1>$1</h1>");

    // Code blocks
    html = html.replace(/```([a-z]*)\n([\s\S]*?)```/gim, "<pre><code>$2</code></pre>");
    // Inline code
    html = html.replace(/`([^`]+)`/g, "<code>$1</code>");

    // Bold, Italic
    html = html.replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>");
    html = html.replace(/\*(.*?)\*/g, "<em>$1</em>");

    // Lists
    html = html.replace(/^\s*[-*]\s+(.*$)/gim, "<li>$1</li>");
    html = html.replace(/(<li>.*<\/li>)/gim, "<ul>$1</ul>");

    // Paragraphs / line breaks
    html = html.replace(/\n\n/g, "<br/><br/>");
    html = html.replace(/\n/g, "<br/>");

    return html;
  }

  escapeHtml(str) {
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  scrollToBottom() {
    this.dom.chatContainer.scrollTop = this.dom.chatContainer.scrollHeight;
  }

  showToast(msg) {
    const existing = document.querySelector(".toast-msg");
    if (existing) existing.remove();

    const toast = document.createElement("div");
    toast.className = "toast-msg";
    toast.textContent = msg;

    document.body.appendChild(toast);
    setTimeout(() => {
      toast.style.opacity = "0";
      toast.style.transform = "translate(-50%, 10px)";
      toast.style.transition = "all 0.2s ease";
      setTimeout(() => toast.remove(), 250);
    }, 2500);
  }
}

// Khởi chạy ứng dụng khi DOM tải xong
document.addEventListener("DOMContentLoaded", () => {
  const app = new AppController();
  app.init();
});
