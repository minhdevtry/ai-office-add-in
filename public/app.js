/**
 * ORI AI — Microsoft Word AI Agent Controller (Professional 3.0)
 * AI Agent Coding IDE Paradigm for Microsoft Word Documents
 */

/* global Office */

// Safe interceptor for window.confirm & window.alert in Office Add-in iframe
if (typeof window !== "undefined") {
  try {
    window.confirm = (msg) => {
      console.warn("[Ori AI] window.confirm intercepted:", msg);
      return false;
    };
    window.alert = (msg) => {
      console.warn("[Ori AI] window.alert intercepted:", msg);
    };
  } catch (_) {}
}

// Filter browser warning spam from third-party scripts (Office SDK)
if (typeof window !== "undefined" && typeof window.console !== "undefined") {
  const _origWarn = window.console.warn.bind(window.console);
  const _origError = window.console.error.bind(window.console);
  const _filter = (args) => {
    const s = args.map((a) => (typeof a === "string" ? a : "")).join(" ");
    if (/Tracking Prevention blocked access to storage/i.test(s)) return;
    if (/Tracking Prevention blocked storage/i.test(s)) return;
    return false;
  };
  window.console.warn = (...args) => {
    if (_filter(args) === undefined) return;
    _origWarn(...args);
  };
  window.console.error = (...args) => {
    if (_filter(args) === undefined) return;
    _origError(...args);
  };
}

// Detect localStorage restriction
if (typeof window !== "undefined" && !window.__ORI_NO_LS__) {
  try {
    const _t = "__ori_ls_probe__";
    window.localStorage.setItem(_t, "1");
    window.localStorage.removeItem(_t);
  } catch (_) {
    window.__ORI_NO_LS__ = true;
    console.warn("[Ori AI] localStorage bị hạn chế — dùng memory fallback.");
  }
}

import { wordBridge } from "./modules/word-bridge.js?v=3.0.0";
import { docState } from "./modules/doc-state.js?v=3.0.0";
import { sseClient } from "./modules/sse-client.js?v=3.0.0";
import { wsClient } from "./modules/ws-client.js?v=3.0.0";
import { licenseClient } from "./modules/license-client.js?v=3.0.0";
import { getOrCreateDeviceId } from "./modules/device-fingerprint.js?v=3.0.0";
import { computeWordDiff } from "./modules/diff.js?v=3.0.0";
import { saveSession, loadLatestSession, listSessions, deleteSession, newThreadId } from "./modules/session-store.js?v=3.0.0";
import { getAllStreamConfigs, saveStreamProxyConfig, saveCorsProxyConfig, saveProviderKeyConfig } from "./modules/stream-factory.js?v=3.0.0";

const DEFAULT_SKILLS = [
  {
    id: "grammar",
    icon: "🔍",
    label: "Soát chính tả & dấu",
    tip: "Tìm và sửa lỗi chính tả tiếng Việt trong toàn bài",
    prompt: "Hãy rà soát toàn bộ tài liệu Word đang mở, tìm tất cả các lỗi chính tả, lỗi gõ dấu tiếng Việt và ngữ pháp. Với mỗi lỗi tìm thấy, hãy xuất marker đề xuất sửa theo cú pháp:\n[Góp ý: \"câu hoặc cụm từ gốc chứa lỗi\" -> \"câu hoặc cụm từ đã sửa đúng\"]\nSau đó nêu tóm tắt ngắn gọn các lỗi đã tìm thấy."
  },
  {
    id: "polish",
    icon: "✍️",
    label: "Viết lại mượt mà",
    tip: "Chỉnh sửa câu văn trôi chảy, tự nhiên và lôi cuốn hơn",
    prompt: "Hãy đọc toàn bộ tài liệu Word và viết lại các đoạn văn còn thô hoặc lủng củng để câu từ trở nên trôi chảy, tự nhiên và thoát ý hơn nhưng giữ nguyên nội dung cốt lõi. Với mỗi đoạn cần sửa, hãy xuất marker theo cú pháp:\n[Góp ý: \"đoạn văn gốc cần viết lại\" -> \"đoạn văn viết lại mượt mà\"]"
  },
  {
    id: "formal",
    icon: "👔",
    label: "Văn phong công vụ",
    tip: "Chuẩn hóa sang văn phong hành chính, trang trọng",
    prompt: "Hãy chuyển đổi các câu từ trong bài viết sang văn phong hành chính, trang trọng và chuẩn mực công vụ. Với mỗi câu cần chỉnh sửa, hãy xuất marker:\n[Góp ý: \"câu gốc thường ngày\" -> \"câu chuẩn mực hành chính\"]"
  },
  {
    id: "cut_fluff",
    icon: "✂️",
    label: "Gọt từ thừa & Sáo rỗng",
    tip: "Rút ngắn 20-30% câu từ, loại bỏ sáo rỗng",
    prompt: "Hãy rà soát tài liệu, loại bỏ tất cả các từ ngữ sáo rỗng, từ đệm vô nghĩa, cô đọng câu văn để giảm 20-30% độ dài nhưng giữ trọn ý nghĩa. Xuất marker cho các điểm sửa:\n[Góp ý: \"đoạn văn gốc dài dòng\" -> \"đoạn văn đã cô đọng\"]"
  },
  {
    id: "ai_proof",
    icon: "🕵️",
    label: "Sửa văn phong AI",
    tip: "Loại bỏ văn mẫu sáo rỗng của AI",
    prompt: "Hãy rà soát và loại bỏ các cấu trúc văn mẫu sáo rỗng thường thấy của AI. Sửa lại để câu từ chân thực, tự nhiên như con người viết:\n[Góp ý: \"câu văn mẫu AI\" -> \"câu văn tự nhiên con người\"]"
  },
  {
    id: "summary",
    icon: "📝",
    label: "Tóm tắt ý chính",
    tip: "Trích xuất các luận điểm cốt lõi gạch đầu dòng",
    prompt: "Hãy tóm tắt các luận điểm cốt lõi của toàn bộ tài liệu dưới dạng các gạch đầu dòng rõ ràng, súc tích và có cấu trúc mạch lạc."
  },
  {
    id: "comment_review",
    icon: "💬",
    label: "Đánh giá biên tập",
    tip: "Nhận xét ưu/nhược điểm toàn diện",
    prompt: "Đóng vai trò là một biên tập viên chuyên nghiệp, hãy đánh giá toàn diện bài viết (bố cục, luận điểm, văn phong). Với mỗi đoạn cần góp ý cụ thể, hãy tạo marker:\n[Góp ý: \"câu văn cần góp ý\" -> \"gợi ý chỉnh sửa hoặc viết lại\"]"
  }
];

class AppController {
  constructor() {
    this.state = {
      messages: [],
      isStreaming: false,
      skills: [...DEFAULT_SKILLS],
      theme: "light",
      config: {
        effort: "medium", // low | medium | high | xhigh
        houseVoice: "",
      },
    };

    this.dom = {};
    this._threadId = newThreadId();
    this._saveSessionDebounced = null;
    this._slashFilteredSkills = [];
    this._slashSelectedIndex = 0;
    this._slashMenuOpen = false;
  }

  async init() {
    this.cacheDom();
    this.initTheme();
    this.initOffice();
    this.bindEvents();
    this.initEffortControls();
    this.initSettingsTabs();

    await this.loadSkills();
    this.renderSlashMenuItems();
    this.loadState();
    this.refreshDocInsights();

    // License Check
    const activated = await this.initLicense();
    if (activated) {
      this.initWebSocket();
      this._loadSessionForCurrentDoc().catch((e) => console.warn("[Ori AI] session restore:", e));
    }
  }

  cacheDom() {
    this.dom = {
      // Header
      quickEffortSelect: document.getElementById("quickEffortSelect"),
      btnToggleTheme: document.getElementById("btnToggleTheme"),
      btnSessions: document.getElementById("btnSessions"),
      btnSettings: document.getElementById("btnSettings"),

      // Main Canvas
      chatContainer: document.getElementById("chatContainer"),
      emptyHint: document.getElementById("emptyHint"),
      docInsightsBar: document.getElementById("docInsightsBar"),
      docInsightsText: document.getElementById("docInsightsText"),

      // Quick Skills & Composer
      quickSkillsBar: document.getElementById("quickSkillsBar"),
      slashMenu: document.getElementById("slashMenu"),
      slashMenuList: document.getElementById("slashMenuList"),
      chatInput: document.getElementById("chatInput"),
      btnSend: document.getElementById("btnSend"),

      // Settings Modal
      settingsModal: document.getElementById("settingsModal"),
      btnCloseSettings: document.getElementById("btnCloseSettings"),
      btnSaveSettings: document.getElementById("btnSaveSettings"),
      cfgHouseVoice: document.getElementById("cfgHouseVoice"),
      cfgEffort: document.getElementById("cfgEffort"),
      btnClearChatModal: document.getElementById("btnClearChatModal"),
      cfgStreamProxyEnabled: document.getElementById("cfgStreamProxyEnabled"),
      cfgStreamProxyUrl: document.getElementById("cfgStreamProxyUrl"),
      cfgStreamProxyToken: document.getElementById("cfgStreamProxyToken"),
      cfgCorsProxyEnabled: document.getElementById("cfgCorsProxyEnabled"),
      cfgCorsProxyUrl: document.getElementById("cfgCorsProxyUrl"),
      cfgDirectEnabled: document.getElementById("cfgDirectEnabled"),
      cfgDirectBaseURL: document.getElementById("cfgDirectBaseURL"),
      cfgDirectApiKey: document.getElementById("cfgDirectApiKey"),
      licEmailDisplay: document.getElementById("licEmailDisplay"),
      licDeviceDisplay: document.getElementById("licDeviceDisplay"),

      // Sessions Modal
      sessionsModal: document.getElementById("sessionsModal"),
      sessionsList: document.getElementById("sessionsList"),
      btnCloseSessions: document.getElementById("btnCloseSessions"),

      // License Modal
      licenseModal: document.getElementById("licenseModal"),
      licEmail: document.getElementById("licEmail"),
      licKey: document.getElementById("licKey"),
      licStatus: document.getElementById("licStatus"),
      btnActivate: document.getElementById("btnActivate"),
    };
  }

  initTheme() {
    const savedTheme = docState.getLocal("ai_word.theme", "light");
    this.setTheme(savedTheme);
  }

  setTheme(theme) {
    this.state.theme = theme;
    document.documentElement.setAttribute("data-theme", theme);
    docState.setLocal("ai_word.theme", theme);
    if (this.dom.btnToggleTheme) {
      this.dom.btnToggleTheme.textContent = theme === "dark" ? "☀️" : "🌙";
      this.dom.btnToggleTheme.title = theme === "dark" ? "Chuyển sang Giao diện Sáng" : "Chuyển sang Giao diện Tối";
    }
  }

  toggleTheme() {
    const next = this.state.theme === "dark" ? "light" : "dark";
    this.setTheme(next);
  }

  initEffortControls() {
    const saved = docState.loadEffort(this.state.config.effort);
    this.state.config.effort = saved;

    if (this.dom.quickEffortSelect) {
      this.dom.quickEffortSelect.value = saved;
      this.dom.quickEffortSelect.addEventListener("change", () => {
        const effort = this.dom.quickEffortSelect.value;
        this.state.config.effort = effort;
        docState.saveEffort(effort);
        if (this.dom.cfgEffort) this.dom.cfgEffort.value = effort;
        this.showToast(`🧠 Mức độ suy luận: ${this.dom.quickEffortSelect.options[this.dom.quickEffortSelect.selectedIndex].text}`);
      });
    }

    if (this.dom.cfgEffort) {
      this.dom.cfgEffort.value = saved;
      this.dom.cfgEffort.addEventListener("change", () => {
        const effort = this.dom.cfgEffort.value;
        this.state.config.effort = effort;
        docState.saveEffort(effort);
        if (this.dom.quickEffortSelect) this.dom.quickEffortSelect.value = effort;
      });
    }
  }

  initSettingsTabs() {
    const tabBtns = document.querySelectorAll(".settings-tabs .tab-btn");
    tabBtns.forEach((btn) => {
      btn.addEventListener("click", () => {
        const targetId = btn.dataset.tab;
        tabBtns.forEach((b) => b.classList.remove("active"));
        btn.classList.add("active");

        const panes = document.querySelectorAll(".modal-body .tab-pane");
        panes.forEach((p) => p.classList.remove("active"));
        const targetPane = document.getElementById(targetId);
        if (targetPane) targetPane.classList.add("active");
      });
    });
  }

  bindEvents() {
    this.dom.btnToggleTheme.addEventListener("click", () => this.toggleTheme());

    // Send / Stop action
    this.dom.btnSend.addEventListener("click", () => this.handleSendOrStop());
    this.dom.chatInput.addEventListener("keydown", (e) => {
      if (this._slashMenuOpen) {
        if (e.key === "ArrowDown") {
          e.preventDefault();
          this.navigateSlashMenu(1);
          return;
        }
        if (e.key === "ArrowUp") {
          e.preventDefault();
          this.navigateSlashMenu(-1);
          return;
        }
        if (e.key === "Enter" || e.key === "Tab") {
          e.preventDefault();
          this.selectCurrentSlashItem();
          return;
        }
        if (e.key === "Escape") {
          e.preventDefault();
          this.closeSlashMenu();
          return;
        }
      }

      if (e.key === "Enter" && !e.shiftKey) {
        e.preventDefault();
        this.handleSendOrStop();
      }
    });

    // Auto-expand textarea & slash trigger
    this.dom.chatInput.addEventListener("input", () => {
      this.dom.chatInput.style.height = "auto";
      this.dom.chatInput.style.height = Math.min(this.dom.chatInput.scrollHeight, 110) + "px";

      const val = this.dom.chatInput.value;
      if (val.startsWith("/")) {
        this.openSlashMenu(val.slice(1).toLowerCase());
      } else {
        this.closeSlashMenu();
      }
    });

    document.addEventListener("click", (e) => {
      if (!e.target.closest(".input-box-wrapper")) {
        this.closeSlashMenu();
      }
    });

    // Starter Cards delegation
    this.dom.emptyHint.addEventListener("click", (e) => {
      const card = e.target.closest(".starter-card");
      if (card) {
        const skillId = card.dataset.skill;
        const skill = this.state.skills.find((s) => s.id === skillId);
        if (skill) this.triggerSkill(skill);
      }
    });

    // Quick Skills Bar delegation
    if (this.dom.quickSkillsBar) {
      this.dom.quickSkillsBar.addEventListener("click", (e) => {
        const btn = e.target.closest(".quick-skill-btn");
        if (btn) {
          const skillId = btn.dataset.skill;
          const skill = this.state.skills.find((s) => s.id === skillId);
          if (skill) this.triggerSkill(skill);
        }
      });
    }

    // Modals
    this.dom.btnSettings.addEventListener("click", () => this.openSettings());
    this.dom.btnCloseSettings.addEventListener("click", () => this.closeSettings());
    this.dom.btnSaveSettings.addEventListener("click", () => this.saveSettings());
    if (this.dom.settingsModal) {
      this.dom.settingsModal.addEventListener("click", (e) => {
        if (e.target === this.dom.settingsModal) this.closeSettings();
      });
    }

    if (this.dom.btnClearChatModal) {
      this.dom.btnClearChatModal.addEventListener("click", () => {
        this.clearChatHistory();
        this.closeSettings();
      });
    }

    this.dom.btnSessions.addEventListener("click", () => this.openSessions());
    if (this.dom.btnCloseSessions) {
      this.dom.btnCloseSessions.addEventListener("click", () => this.closeSessions());
    }
    if (this.dom.sessionsModal) {
      this.dom.sessionsModal.addEventListener("click", (e) => {
        if (e.target === this.dom.sessionsModal) this.closeSessions();
      });
    }

    // License Modal
    if (this.dom.btnActivate) {
      this.dom.btnActivate.addEventListener("click", () => this.handleActivate());
    }
  }

  async refreshDocInsights() {
    try {
      const text = await wordBridge.getFullDocumentText(30000);
      if (text && text.trim()) {
        const stats = wordBridge.getTextAnalytics(text);
        if (this.dom.docInsightsBar && this.dom.docInsightsText) {
          this.dom.docInsightsBar.style.display = "flex";
          this.dom.docInsightsText.textContent = `Toàn bộ tài liệu: ${stats.words.toLocaleString()} từ • ~${stats.readingTimeMinutes} phút đọc`;
        }
      }
    } catch (_) {}
  }

  async loadSkills() {
    try {
      const res = await fetch("skills.json?v=3.0.0");
      if (res.ok) {
        const data = await res.json();
        if (Array.isArray(data) && data.length > 0) {
          this.state.skills = data;
          return;
        }
      }
    } catch (_) {}
    this.state.skills = DEFAULT_SKILLS;
  }

  initOffice() {
    if (typeof Office !== "undefined" && Office.onReady) {
      Office.onReady((info) => {
        if (info && info.host === Office.HostType.Word) {
          this.refreshDocInsights();
        }
      }).catch((err) => {
        console.warn("[Ori AI] Office.onReady error:", err);
      });
    }
  }

  initWebSocket() {
    wsClient.connect();
  }

  // ─── LICENSE LIFECYCLE ────────────────────────────────────────────────
  async initLicense() {
    getOrCreateDeviceId();
    const publicConfig = await licenseClient.fetchPublicConfig();
    if (publicConfig?.aiDefaultEffort) {
      const saved = docState.loadEffort(null);
      if (!saved) {
        this.state.config.effort = publicConfig.aiDefaultEffort;
        docState.saveEffort(publicConfig.aiDefaultEffort);
      }
    }

    if (publicConfig && publicConfig.licenseRequired === false) {
      return true;
    }

    if (licenseClient.isActivated()) {
      licenseClient.startHeartbeat();
      licenseClient.onInvalidate(() => this.showLicenseModal(true));
      return true;
    }

    this.showLicenseModal(false);
    return false;
  }

  showLicenseModal(showStatus) {
    if (!this.dom.licenseModal) return;
    this.dom.licenseModal.classList.add("open");
    if (showStatus && this.dom.licStatus) {
      this.showLicenseStatus("error", "Bản quyền đã hết hạn. Vui lòng kích hoạt lại.");
    }
  }

  hideLicenseModal() {
    if (!this.dom.licenseModal) return;
    this.dom.licenseModal.classList.remove("open");
  }

  showLicenseStatus(type, msg) {
    if (!this.dom.licStatus) return;
    this.dom.licStatus.className = `license-status show ${type}`;
    this.dom.licStatus.textContent = msg;
  }

  async handleActivate() {
    const email = this.dom.licEmail.value.trim();
    const licenseKey = this.dom.licKey.value.trim();

    if (!email || !licenseKey) {
      this.showLicenseStatus("error", "Vui lòng nhập đầy đủ email và License Key.");
      return;
    }

    this.dom.btnActivate.disabled = true;
    this.dom.btnActivate.textContent = "Đang kích hoạt...";

    try {
      const deviceName = navigator.userAgent.includes("Word") ? "Word Desktop" : "Web";
      await licenseClient.activate({ email, licenseKey, deviceName });
      this.showLicenseStatus("success", "✅ Kích hoạt thành công!");

      licenseClient.startHeartbeat();
      licenseClient.onInvalidate(() => this.showLicenseModal(true));

      setTimeout(() => {
        this.hideLicenseModal();
        this.initWebSocket();
        this.loadState();
        this.refreshDocInsights();
        this.showToast("✅ Kích hoạt bản quyền thành công!");
      }, 700);
    } catch (err) {
      this.showLicenseStatus("error", err.message || "Kích hoạt thất bại. Vui lòng thử lại.");
    } finally {
      this.dom.btnActivate.disabled = false;
      this.dom.btnActivate.textContent = "Kích hoạt";
    }
  }

  // ─── STATE & SESSIONS ─────────────────────────────────────────────────
  loadState() {
    this.state.messages = docState.loadHistory();
    this.state.config.houseVoice = docState.loadHouseVoice();
    this.state.config.effort = docState.loadEffort(this.state.config.effort);

    if (this.state.messages.length > 0) {
      this.dom.emptyHint.style.display = "none";
      for (const msg of this.state.messages) {
        this.renderMessageBubble(msg.role, msg.displayContent || msg.content, {
          thinking: msg.thinking,
          persist: false,
        });
      }
    }
  }

  async _persistSession() {
    try {
      const docPath = (await wordBridge.getDocumentPathSafe?.()) || "unknown";
      const systemPrompt = docState.loadSystemPrompt();
      await saveSession({
        threadId: this._threadId,
        docPath,
        messages: this.state.messages,
        systemPrompt,
        config: this.state.config,
      });
    } catch (err) {
      console.warn("[Ori AI] persistSession error:", err.message);
    }
  }

  async _loadSessionForCurrentDoc() {
    try {
      const docPath = (await wordBridge.getDocumentPathSafe?.()) || "unknown";
      const session = await loadLatestSession(docPath);
      if (!session) return false;
      this._threadId = session.threadId;
      this.state.messages = session.messages || [];
      this.state.config = { ...this.state.config, ...(session.config || {}) };
      if (this.state.messages.length > 0) {
        this.dom.emptyHint.style.display = "none";
        this.dom.chatContainer.innerHTML = "";
        for (const msg of this.state.messages) {
          this.renderMessageBubble(msg.role, msg.displayContent || msg.content, {
            thinking: msg.thinking,
            persist: false,
          });
        }
      }
      return true;
    } catch (err) {
      return false;
    }
  }

  _schedulePersistSession() {
    if (this._saveSessionDebounced) clearTimeout(this._saveSessionDebounced);
    this._saveSessionDebounced = setTimeout(() => this._persistSession(), 800);
  }

  clearChatHistory() {
    docState.clearHistory();
    this.state.messages = [];
    this.dom.chatContainer.innerHTML = "";
    this.dom.chatContainer.appendChild(this.dom.emptyHint);
    this.dom.emptyHint.style.display = "flex";
    this.refreshDocInsights();
    this.showToast("🗑️ Đã xóa lịch sử chat.");
  }

  // ─── CHAT & STREAMING (FULL-DOCUMENT CONTEXT) ─────────────────────────
  async triggerSkill(skill) {
    let fullDocText = "";
    try {
      fullDocText = await wordBridge.getFullDocumentText(50000);
    } catch (_) {}

    let prompt = skill.prompt;
    if (fullDocText && fullDocText.trim()) {
      prompt = `${skill.prompt}\n\n[Tài liệu Word đang mở (${wordBridge.getTextAnalytics(fullDocText).words} từ)]:\n"""\n${fullDocText.trim()}\n"""`;
    }

    this.sendMessage(prompt, {
      displayPrompt: `${skill.icon} ${skill.label}`,
      originalText: fullDocText || "",
    });
  }

  async handleSendOrStop() {
    if (this.state.isStreaming) {
      sseClient.abort();
      this.setStreamingState(false);
      return;
    }

    const userInput = this.dom.chatInput.value.trim();
    if (!userInput) return;

    this.dom.chatInput.value = "";
    this.dom.chatInput.style.height = "auto";
    this.closeSlashMenu();

    let fullDocText = "";
    try {
      fullDocText = await wordBridge.getFullDocumentText(50000);
    } catch (_) {}

    let fullPrompt = userInput;
    if (fullDocText && fullDocText.trim()) {
      fullPrompt = `${userInput}\n\n[Tài liệu Word đang mở (${wordBridge.getTextAnalytics(fullDocText).words} từ)]:\n"""\n${fullDocText.trim()}\n"""`;
    }

    this.sendMessage(fullPrompt, {
      displayPrompt: userInput,
      originalText: fullDocText || "",
    });
  }

  async sendMessage(promptText, meta = {}) {
    this.dom.emptyHint.style.display = "none";

    // 1. Render User Message
    const userDisplay = meta.displayPrompt || promptText;
    this.renderMessageBubble("user", userDisplay);
    this.state.messages.push({ role: "user", content: promptText, displayContent: userDisplay });
    docState.saveHistory(this.state.messages);
    this._schedulePersistSession();

    // 2. Create Assistant Bubble
    const assistantBubble = this.createAssistantBubble();
    this.setStreamingState(true);

    // 3. System Prompt & House Voice
    let systemPrompt = docState.loadSystemPrompt();
    const houseVoice = this.state.config.houseVoice || docState.loadHouseVoice();
    if (houseVoice && houseVoice.trim()) {
      systemPrompt += `\n\n<personal_writing_style>\n${houseVoice.trim()}\nQuy tắc: Luôn mô phỏng đúng văn phong, nhịp điệu và cách dùng từ của đoạn mẫu trên.\n</personal_writing_style>`;
    }

    // 4. Real-time Comment Marker Spawning
    const spawnedMarkers = new Set();
    const handleNewMarkers = async (fullText) => {
      const markers = this.extractCommentMarkers(fullText);
      for (const marker of markers) {
        const key = `${marker.anchor}|||${marker.suggestion}`;
        if (spawnedMarkers.has(key)) continue;
        spawnedMarkers.add(key);

        // Ghim Comment lề trang vào Word với prefix [Ori Agent]
        try {
          await wordBridge.insertComment(
            `Góp ý: "${marker.anchor}"\n→ Sửa thành: "${marker.suggestion}"`,
            marker.anchor,
            { authorTag: true }
          );
        } catch (err) {
          console.warn("[Ori AI] spawn comment error:", err.message);
        }
      }
    };

    try {
      await sseClient.streamChat({
        messages: this.state.messages,
        systemPrompt,
        effort: this.state.config.effort,

        onThinking: (chunk, fullThinking) => {
          assistantBubble.updateThinking(fullThinking);
          this.scrollToBottom();
        },

        onDelta: (chunk, fullText) => {
          assistantBubble.updateContent(fullText, true);
          handleNewMarkers(fullText);
          this.scrollToBottom();
        },

        onDone: (fullText, fullThinking) => {
          this.setStreamingState(false);
          handleNewMarkers(fullText);
          assistantBubble.finalize(fullText, fullThinking, meta);

          this.state.messages.push({
            role: "assistant",
            content: fullText,
            thinking: fullThinking,
          });
          docState.saveHistory(this.state.messages);
          this._schedulePersistSession();
          this.refreshDocInsights();
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

  // ─── MARKER PARSER (Interactive Patch Cards) ──────────────────────────
  extractCommentMarkers(text) {
    if (!text || typeof text !== "string") return [];
    const out = [];
    const regex = /\[(?:Góp ý|Comment|Nhận xét|Feedback):\s*"([^"]{3,400})"\s*(?:->|:|➔)\s*([^\]]+)\]/gim;
    let m;
    while ((m = regex.exec(text)) !== null) {
      let sug = m[2].trim();
      sug = sug.replace(/^["']|["']$/g, "").trim();
      out.push({
        anchor: m[1].trim(),
        suggestion: sug,
      });
    }
    return out;
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

  // ─── RENDERING ASSISTANT BUBBLE & PATCH CARDS ─────────────────────────
  renderMessageBubble(role, content, { thinking = "", persist = true, meta = {} } = {}) {
    const bubble = document.createElement("div");
    bubble.className = `message-bubble ${role}`;

    if (role === "user") {
      bubble.textContent = String(content).replace(/^[\u{1F300}-\u{1FAFF}\u{2600}-\u{27BF}]\s*/u, "").trim();
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

      // Render Diff Patch Cards
      const markers = this.extractCommentMarkers(content);
      if (markers.length > 0) {
        this.renderPatchSuite(bubble, markers);
      }

      // Render General Actions
      this.attachGeneralActions(bubble, content);
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
        contentDiv.innerHTML = this.formatMarkdown(text) + (isStreaming ? '<span class="streaming-cursor"></span>' : "");
      },

      finalize: (fullText, fullThinking) => {
        contentDiv.innerHTML = this.formatMarkdown(fullText);

        // Render Patch Cards
        const markers = this.extractCommentMarkers(fullText);
        if (markers.length > 0) {
          this.renderPatchSuite(bubble, markers);
        }

        // Render General Actions
        this.attachGeneralActions(bubble, fullText);
      },

      showError: (errText) => {
        contentDiv.innerHTML += `<div style="color: var(--danger); margin-top: 6px; font-weight: 500;">⚠️ ${this.escapeHtml(errText)}</div>`;
      },
    };
  }

  // ─── PATCH SUITE COMPONENT (AI CODING AGENT PARADIGM) ─────────────────
  renderPatchSuite(bubble, markers) {
    const container = document.createElement("div");
    container.className = "patch-suite-container";

    // Multi-patch toolbar (khi có >= 2 đề xuất)
    if (markers.length >= 2) {
      const toolbar = document.createElement("div");
      toolbar.className = "patch-toolbar";
      toolbar.innerHTML = `
        <span class="patch-count-badge">⚡ ${markers.length} đề xuất chỉnh sửa</span>
        <div class="patch-toolbar-actions">
          <button class="btn-accept-all">✓ Áp dụng tất cả</button>
          <button class="btn-reject-all">✕ Bỏ qua tất cả</button>
        </div>
      `;

      toolbar.querySelector(".btn-accept-all").addEventListener("click", async () => {
        toolbar.querySelector(".btn-accept-all").disabled = true;
        toolbar.querySelector(".btn-accept-all").textContent = "Đang áp dụng...";
        for (const m of markers) {
          try {
            await wordBridge.acceptComment({ anchorText: m.anchor, replacementText: m.suggestion });
          } catch (_) {}
        }
        container.querySelectorAll(".patch-card").forEach((card) => {
          card.className = "patch-card card-accepted";
          const statusTag = card.querySelector(".patch-status-tag");
          if (statusTag) {
            statusTag.className = "patch-status-tag accepted";
            statusTag.textContent = "✅ Đã áp dụng";
          }
          card.querySelectorAll(".patch-actions button").forEach((b) => (b.disabled = true));
        });
        toolbar.remove();
        this.showToast("✅ Đã áp dụng tất cả sửa đổi vào Word!");
        this.refreshDocInsights();
      });

      toolbar.querySelector(".btn-reject-all").addEventListener("click", async () => {
        toolbar.remove();
        for (const m of markers) {
          try {
            await wordBridge.deleteComment({ anchorText: m.anchor });
          } catch (_) {}
        }
        container.querySelectorAll(".patch-card").forEach((card) => {
          card.className = "patch-card card-rejected";
          const statusTag = card.querySelector(".patch-status-tag");
          if (statusTag) {
            statusTag.className = "patch-status-tag rejected";
            statusTag.textContent = "✕ Đã bỏ qua";
          }
          card.querySelectorAll(".patch-actions button").forEach((b) => (b.disabled = true));
        });
        this.showToast("🗑️ Đã bỏ qua tất cả đề xuất.");
      });

      container.appendChild(toolbar);
    }

    // Render từng Card sửa đổi
    markers.forEach((marker, idx) => {
      const card = document.createElement("div");
      card.className = "patch-card";

      // Compute word diff for visual preview
      const diffChunks = computeWordDiff(marker.anchor, marker.suggestion);
      const diffHtml = diffChunks
        .map((c) => {
          const val = this.escapeHtml(c.value);
          if (c.type === "del") return `<del class="diff-del">${val}</del>`;
          if (c.type === "ins") return `<ins class="diff-ins">${val}</ins>`;
          return val;
        })
        .join("");

      card.innerHTML = `
        <div class="patch-card-header">
          <span class="patch-badge">📝 Sửa đổi #${idx + 1}</span>
          <span class="patch-status-tag">Chờ duyệt</span>
        </div>
        <div class="patch-diff-body">${diffHtml}</div>
        <div class="patch-actions">
          <button class="btn-patch-accept">✓ Áp dụng</button>
          <button class="btn-patch-reject">✕ Bỏ qua</button>
        </div>
      `;

      // Accept handler: replace in Word & delete comment
      card.querySelector(".btn-patch-accept").addEventListener("click", async () => {
        card.querySelectorAll("button").forEach((b) => (b.disabled = true));
        try {
          await wordBridge.acceptComment({
            anchorText: marker.anchor,
            replacementText: marker.suggestion,
          });
          card.className = "patch-card card-accepted";
          const statusTag = card.querySelector(".patch-status-tag");
          statusTag.className = "patch-status-tag accepted";
          statusTag.textContent = "✅ Đã áp dụng";
          this.showToast("✅ Đã áp dụng sửa đổi vào Word!");
          this.refreshDocInsights();
        } catch (err) {
          card.querySelectorAll("button").forEach((b) => (b.disabled = false));
          this.showToast(`⚠️ Lỗi: ${err.message}`);
        }
      });

      // Reject handler: delete comment from Word
      card.querySelector(".btn-patch-reject").addEventListener("click", async () => {
        card.querySelectorAll("button").forEach((b) => (b.disabled = true));
        try {
          await wordBridge.deleteComment({ anchorText: marker.anchor });
        } catch (_) {}
        card.className = "patch-card card-rejected";
        const statusTag = card.querySelector(".patch-status-tag");
        statusTag.className = "patch-status-tag rejected";
        statusTag.textContent = "✕ Đã bỏ qua";
        this.showToast("🗑️ Đã bỏ qua sửa đổi.");
      });

      container.appendChild(card);
    });

    bubble.appendChild(container);
  }

  // ─── GENERAL ACTIONS (INSERT AT CURSOR & COPY) ────────────────────────
  attachGeneralActions(bubble, text) {
    const actionsBar = document.createElement("div");
    actionsBar.className = "msg-actions";
    const cleanText = this.extractExecutableText(text);

    // 1. Chèn vào con trỏ
    const btnInsert = document.createElement("button");
    btnInsert.className = "action-btn";
    btnInsert.innerHTML = `<span>➕</span> Chèn vào con trỏ`;
    btnInsert.addEventListener("click", async () => {
      try {
        await wordBridge.insertAtCursor(cleanText);
        this.showToast("✅ Đã chèn nội dung vào vị trí con trỏ Word!");
        this.refreshDocInsights();
      } catch (err) {
        this.showToast(`⚠️ Lỗi chèn: ${err.message}`);
      }
    });
    actionsBar.appendChild(btnInsert);

    // 2. Sao chép
    const btnCopy = document.createElement("button");
    btnCopy.className = "action-btn";
    btnCopy.innerHTML = `<span>📋</span> Sao chép`;
    btnCopy.addEventListener("click", () => {
      navigator.clipboard.writeText(cleanText);
      this.showToast("📋 Đã sao chép vào bộ nhớ tạm!");
    });
    actionsBar.appendChild(btnCopy);

    bubble.appendChild(actionsBar);
  }

  extractExecutableText(text) {
    if (!text || typeof text !== "string") return "";

    // 1. Văn bản giữa 2 dòng ---
    const delimiterMatch = text.match(/(?:^|\n)\s*---\s*\n([\s\S]*?)\n\s*---\s*(?:\n|$)/);
    if (delimiterMatch && delimiterMatch[1] && delimiterMatch[1].trim()) {
      return delimiterMatch[1].trim();
    }

    // 2. Code block ```
    const codeBlockMatch = text.match(/```(?:[a-z]*)\n([\s\S]*?)```/i);
    if (codeBlockMatch && codeBlockMatch[1] && codeBlockMatch[1].trim()) {
      return codeBlockMatch[1].trim();
    }

    // 3. Fallback: loại bỏ marker tags
    return text.replace(/\[(?:Góp ý|Comment|Nhận xét|Feedback):\s*"[^"]+"\s*(?:->|:|➔)\s*[^\]]+\]/gim, "").trim();
  }

  // ─── SLASH COMMANDS MENU ──────────────────────────────────────────────
  openSlashMenu(filterQuery = "") {
    if (!this.dom.slashMenu || !this.dom.slashMenuList) return;

    const filtered = this.state.skills.filter(
      (s) => s.label.toLowerCase().includes(filterQuery) || s.id.toLowerCase().includes(filterQuery)
    );

    if (filtered.length === 0) {
      this.closeSlashMenu();
      return;
    }

    this._slashFilteredSkills = filtered;
    this._slashSelectedIndex = 0;
    this._slashMenuOpen = true;
    this.dom.slashMenu.style.display = "flex";
    this.renderSlashMenuItems();
  }

  renderSlashMenuItems() {
    if (!this.dom.slashMenuList) return;
    this.dom.slashMenuList.innerHTML = "";

    this._slashFilteredSkills.forEach((skill, idx) => {
      const item = document.createElement("div");
      item.className = `slash-item ${idx === this._slashSelectedIndex ? "selected" : ""}`;
      item.innerHTML = `
        <span class="slash-item-icon">${skill.icon}</span>
        <span class="slash-item-label">${this.escapeHtml(skill.label)}</span>
        <span class="slash-item-tip">${this.escapeHtml(skill.tip)}</span>
      `;
      item.addEventListener("click", () => this.executeSlashSkill(skill));
      this.dom.slashMenuList.appendChild(item);
    });
  }

  navigateSlashMenu(dir) {
    if (!this._slashFilteredSkills.length) return;
    this._slashSelectedIndex = (this._slashSelectedIndex + dir + this._slashFilteredSkills.length) % this._slashFilteredSkills.length;
    this.renderSlashMenuItems();
  }

  selectCurrentSlashItem() {
    if (this._slashFilteredSkills[this._slashSelectedIndex]) {
      this.executeSlashSkill(this._slashFilteredSkills[this._slashSelectedIndex]);
    }
  }

  executeSlashSkill(skill) {
    this.closeSlashMenu();
    this.dom.chatInput.value = "";
    this.dom.chatInput.style.height = "auto";
    this.triggerSkill(skill);
  }

  closeSlashMenu() {
    this._slashMenuOpen = false;
    if (this.dom.slashMenu) this.dom.slashMenu.style.display = "none";
  }

  // ─── SETTINGS MODAL ───────────────────────────────────────────────────
  openSettings() {
    this.dom.cfgHouseVoice.value = this.state.config.houseVoice || "";
    this.dom.cfgEffort.value = this.state.config.effort || "medium";

    const all = getAllStreamConfigs();
    this.dom.cfgStreamProxyEnabled.checked = all.streamProxy.enabled || false;
    this.dom.cfgStreamProxyUrl.value = all.streamProxy.url || "";
    this.dom.cfgStreamProxyToken.value = all.streamProxy.token || "";
    this.dom.cfgCorsProxyEnabled.checked = all.corsProxy.enabled || false;
    this.dom.cfgCorsProxyUrl.value = all.corsProxy.url || "";
    this.dom.cfgDirectEnabled.checked = !!all.providerKey.apiKey;
    this.dom.cfgDirectBaseURL.value = all.providerKey.baseURL || "";
    this.dom.cfgDirectApiKey.value = all.providerKey.apiKey || "";

    const lic = docState.loadLicense();
    if (this.dom.licEmailDisplay) this.dom.licEmailDisplay.textContent = lic?.email || "—";
    if (this.dom.licDeviceDisplay) this.dom.licDeviceDisplay.textContent = lic?.deviceId ? `${lic.deviceId.substring(0, 12)}...` : "—";

    this.dom.settingsModal.classList.add("open");
  }

  closeSettings() {
    this.dom.settingsModal.classList.remove("open");
  }

  saveSettings() {
    this.state.config.houseVoice = this.dom.cfgHouseVoice.value.trim();
    docState.saveHouseVoice(this.state.config.houseVoice);

    const effort = this.dom.cfgEffort.value;
    this.state.config.effort = effort;
    docState.saveEffort(effort);
    if (this.dom.quickEffortSelect) this.dom.quickEffortSelect.value = effort;

    saveStreamProxyConfig({
      enabled: this.dom.cfgStreamProxyEnabled.checked,
      url: this.dom.cfgStreamProxyUrl.value.trim(),
      token: this.dom.cfgStreamProxyToken.value.trim(),
    });
    saveCorsProxyConfig({
      enabled: this.dom.cfgCorsProxyEnabled.checked,
      url: this.dom.cfgCorsProxyUrl.value.trim(),
    });
    saveProviderKeyConfig({
      provider: "openai",
      apiKey: this.dom.cfgDirectApiKey.value.trim(),
      baseURL: this.dom.cfgDirectBaseURL.value.trim(),
    });

    this.closeSettings();
    this.showToast("✅ Đã lưu cài đặt thành công!");
  }

  // ─── SESSIONS MODAL ───────────────────────────────────────────────────
  async openSessions() {
    this.dom.sessionsModal.classList.add("open");
    await this._loadSessionsList();
  }

  closeSessions() {
    this.dom.sessionsModal.classList.remove("open");
  }

  async _loadSessionsList() {
    if (!this.dom.sessionsList) return;
    this.dom.sessionsList.innerHTML = "<p style='font-size:12px; color:var(--text-tertiary);'>Đang tải lịch sử...</p>";

    try {
      const docPath = (await wordBridge.getDocumentPathSafe?.()) || "unknown";
      const sessions = await listSessions(docPath);

      if (!sessions || sessions.length === 0) {
        this.dom.sessionsList.innerHTML = "<p style='font-size:12px; color:var(--text-tertiary); padding:10px 0;'>Chưa có phiên chat nào trước đó.</p>";
        return;
      }

      this.dom.sessionsList.innerHTML = "";
      sessions.forEach((s) => {
        const item = document.createElement("div");
        item.className = "session-item";
        const dateStr = s.updatedAt ? new Date(s.updatedAt).toLocaleString("vi-VN") : "Gần đây";
        const firstUserMsg = s.messages?.find((m) => m.role === "user")?.displayContent || s.messages?.find((m) => m.role === "user")?.content || "Phiên chat";

        item.innerHTML = `
          <div class="session-item-info">
            <span class="session-item-title">${this.escapeHtml(firstUserMsg.slice(0, 38))}</span>
            <span class="session-item-date">${dateStr} • ${s.messages?.length || 0} tin nhắn</span>
          </div>
          <button class="session-item-del" title="Xóa phiên này">🗑️</button>
        `;

        item.querySelector(".session-item-info").addEventListener("click", () => {
          this._restoreSession(s);
          this.closeSessions();
        });

        item.querySelector(".session-item-del").addEventListener("click", async (e) => {
          e.stopPropagation();
          await deleteSession(s.threadId);
          item.remove();
          this.showToast("🗑️ Đã xóa phiên chat.");
        });

        this.dom.sessionsList.appendChild(item);
      });
    } catch (err) {
      this.dom.sessionsList.innerHTML = `<p style='font-size:12px; color:var(--danger);'>Lỗi: ${err.message}</p>`;
    }
  }

  _restoreSession(session) {
    this._threadId = session.threadId;
    this.state.messages = session.messages || [];
    this.state.config = { ...this.state.config, ...(session.config || {}) };
    docState.saveHistory(this.state.messages);

    this.dom.chatContainer.innerHTML = "";
    if (this.state.messages.length > 0) {
      this.dom.emptyHint.style.display = "none";
      for (const msg of this.state.messages) {
        this.renderMessageBubble(msg.role, msg.displayContent || msg.content, {
          thinking: msg.thinking,
          persist: false,
        });
      }
    } else {
      this.dom.chatContainer.appendChild(this.dom.emptyHint);
      this.dom.emptyHint.style.display = "flex";
    }
    this.showToast(`📂 Đã khôi phục ${this.state.messages.length} tin nhắn.`);
  }

  // ─── UTILITIES ────────────────────────────────────────────────────────
  formatMarkdown(text) {
    if (!text || typeof text !== "string") return "";
    let html = this.escapeHtml(text);

    // Format bold & italic
    html = html.replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>");
    html = html.replace(/\*(.*?)\*/g, "<em>$1</em>");

    // Format headers
    html = html.replace(/^### (.*$)/gim, "<h3>$1</h3>");
    html = html.replace(/^## (.*$)/gim, "<h2>$1</h2>");
    html = html.replace(/^# (.*$)/gim, "<h1>$1</h1>");

    // Format lists
    html = html.replace(/^[-*] (.*$)/gim, "<li>$1</li>");
    html = html.replace(/(<li>.*<\/li>\n?)+/g, (m) => `<ul>${m}</ul>`);

    // Format paragraph breaks
    html = html.replace(/\n{2,}/g, "</p><p>");
    html = `<p>${html}</p>`.replace(/<p><(h\d|ul|li)/g, "<$1").replace(/<\/(h\d|ul|li)><\/p>/g, "</$1>");

    // Filter raw markers from main markdown view to keep it clean
    html = html.replace(/\[(?:Góp ý|Comment|Nhận xét|Feedback):\s*&quot;[^&]+&quot;\s*(?:-&gt;|:|➔)\s*[^\]]+\]/gim, "");

    return html;
  }

  escapeHtml(str) {
    if (!str) return "";
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  scrollToBottom() {
    if (this.dom.chatContainer) {
      this.dom.chatContainer.scrollTop = this.dom.chatContainer.scrollHeight;
    }
  }

  showToast(msg) {
    let toast = document.getElementById("oriToast");
    if (!toast) {
      toast = document.createElement("div");
      toast.id = "oriToast";
      toast.className = "toast-notice";
      document.body.appendChild(toast);
    }
    toast.textContent = msg;
    toast.classList.add("show");
    setTimeout(() => toast.classList.remove("show"), 2400);
  }
}

// ─── BOOTSTRAP ────────────────────────────────────────────────────────
const app = new AppController();
if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", () => app.init());
} else {
  app.init();
}
