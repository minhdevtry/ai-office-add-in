/**
 * AI Word Taskpane - Main Application Controller (2026 Premium Edition)
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

// ─── Detect localStorage bị Tracking Prevention chặn (Office iframe) ───
// Browser vẫn log warning mỗi lần gọi localStorage trong third-party context,
// kể cả khi wrap trong try/catch. Detect 1 lần lúc init, set flag global
// để các module skip localStorage và dùng memory fallback.
if (typeof window !== "undefined" && !window.__ORI_NO_LS__) {
  try {
    const _t = "__ori_ls_probe__";
    window.localStorage.setItem(_t, "1");
    window.localStorage.removeItem(_t);
  } catch (_) {
    window.__ORI_NO_LS__ = true;
    console.warn("[Ori AI] localStorage bị chặn (Tracking Prevention) — dùng memory fallback.");
  }
}

import { wordBridge } from "./modules/word-bridge.js?v=2.3.0";
import { docState } from "./modules/doc-state.js?v=2.3.0";
import { sseClient } from "./modules/sse-client.js?v=2.3.0";
import { wsClient } from "./modules/ws-client.js?v=2.3.0";
import { licenseClient } from "./modules/license-client.js?v=2.3.0";
import { getOrCreateDeviceId } from "./modules/device-fingerprint.js?v=2.3.0";
import { computeWordDiff, renderDiffHtml } from "./modules/diff.js?v=2.3.0";
import { saveSession, loadLatestSession, listSessions, deleteSession, newThreadId } from "./modules/session-store.js?v=2.3.0";
import { getAllStreamConfigs, saveStreamProxyConfig, saveCorsProxyConfig, saveProviderKeyConfig } from "./modules/stream-factory.js?v=2.3.0";

const DEFAULT_SKILLS = [
  {
    id: "polish",
    icon: "✍️",
    label: "Viết lại mượt",
    tip: "Chỉnh sửa câu văn trôi chảy, tự nhiên và lôi cuốn hơn",
    prompt: "Hãy viết lại đoạn văn sau cho trôi chảy, tự nhiên, văn phong mạch lạc và lôi cuốn hơn nhưng giữ nguyên nội dung cốt lõi.",
    context: "selection",
    defaultMode: "replace"
  },
  {
    id: "grammar",
    icon: "🔍",
    label: "Sửa chính tả",
    tip: "Tìm và sửa lỗi ngữ pháp, chính tả tiếng Việt",
    prompt: "Hãy tìm và sửa tất cả các lỗi chính tả, lỗi gõ dấu tiếng Việt và ngữ pháp trong đoạn văn sau. Liệt kê rõ những từ đã sửa và đưa ra bản văn đã sửa hoàn chỉnh.",
    context: "selection",
    defaultMode: "replace"
  },
  {
    id: "formal",
    icon: "👔",
    label: "Trang trọng hóa",
    tip: "Chuyển sang văn phong hành chính, công vụ",
    prompt: "Hãy chuyển đổi đoạn văn sau sang văn phong hành chính, trang trọng, chuẩn mực công vụ.",
    context: "selection",
    defaultMode: "replace"
  },
  {
    id: "cut_fluff",
    icon: "✂️",
    label: "Gọt từ thừa",
    tip: "Rút ngắn 20-30% câu từ, bỏ từ sáo rỗng",
    prompt: "Hãy cô đọng đoạn văn sau, loại bỏ các từ ngữ dư thừa, sáo rỗng, giảm 20-30% độ dài nhưng giữ nguyên toàn bộ ý chính.",
    context: "selection",
    defaultMode: "replace"
  },
  {
    id: "summary",
    icon: "📝",
    label: "Tóm tắt ý",
    tip: "Trích xuất các luận điểm cốt lõi gạch đầu dòng",
    prompt: "Hãy tóm tắt các luận điểm cốt lõi của văn bản sau dưới dạng các gạch đầu dòng rõ ràng, súc tích.",
    context: "auto",
    defaultMode: "after"
  },
  {
    id: "translate_en",
    icon: "🌐",
    label: "Dịch Việt ➔ Anh",
    tip: "Dịch chuẩn xác sang tiếng Anh tự nhiên",
    prompt: "Hãy dịch chuẩn xác đoạn văn bản sau sang tiếng Anh tự nhiên, học thuật và chuyên nghiệp.",
    context: "selection",
    defaultMode: "after"
  },
  {
    id: "comment_review",
    icon: "💬",
    label: "Nhận xét lề trang",
    tip: "Đưa ra đánh giá, nhận xét bên lề tài liệu",
    prompt: "Đóng vai trò là một biên tập viên chuyên nghiệp, hãy đưa ra nhận xét, đánh giá ưu/nhược điểm và gợi ý cải thiện cho đoạn văn bản sau.",
    context: "selection",
    defaultMode: "comment"
  },
  {
    id: "ai_proof",
    icon: "🕵️",
    label: "Sửa văn phong AI",
    tip: "Loại bỏ văn mẫu sáo rỗng thường thấy của AI",
    prompt: "Hãy viết lại đoạn văn sau để câu từ nghe chân thực, tự nhiên như con người viết, loại bỏ các cấu trúc sáo rỗng hay lặp lại của AI.",
    context: "selection",
    defaultMode: "replace"
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
        effort: "medium", // off | low | medium | high | xhigh
        houseVoice: "",
      },
    };

    this.dom = {};
    this._threadId = newThreadId();
    this._saveSessionDebounced = null;
    this._slashFilteredSkills = []; // tránh undefined khi init
  }

  async init() {
    this.cacheDom();
    this.initTheme();
    this.initOffice();
    this.bindEvents();
    this.initEffortDropdown();

    await this.loadSkills();
    // renderSkills() bỏ — slash menu (/) vẫn dùng state.skills làm data source
    this.renderSlashMenuItems();

    // License gate: phải kích hoạt trước khi dùng app
    const activated = await this.initLicense();
    if (!activated) {
      // Show modal block, không init phần còn lại
      return;
    }

    this.initWebSocket();
    this.loadState();
    // Thử khôi phục session theo file Word hiện tại (IndexedDB checkpointer)
    this._loadSessionForCurrentDoc().catch((e) => console.warn("[Ori AI] session restore:", e));
  }

  /**
   * Khởi tạo dropdown Effort (trong Settings modal) — bind sự kiện change
   * Lưu vào localStorage (KHÔNG lưu vào document.settings)
   */
  initEffortDropdown() {
    if (!this.dom.cfgEffort) return;
    this.dom.cfgEffort.value = this.state.config.effort;
    this.dom.cfgEffort.addEventListener("change", () => {
      const effort = this.dom.cfgEffort.value;
      this.state.config.effort = effort;
      docState.saveEffort(effort);
    });
  }

  cacheDom() {
    this.dom = {
      statusPill: document.getElementById("statusPill"),
      statusText: document.getElementById("statusText"),
      btnToggleTheme: document.getElementById("btnToggleTheme"),
      btnSettings: document.getElementById("btnSettings"),
      btnSessions: document.getElementById("btnSessions"),
      sessionsModal: document.getElementById("sessionsModal"),
      sessionsList: document.getElementById("sessionsList"),
      btnCloseSessions: document.getElementById("btnCloseSessions"),
      cfgEffort: document.getElementById("cfgEffort"),
      cfgStreamProxyEnabled: document.getElementById("cfgStreamProxyEnabled"),
      cfgStreamProxyUrl: document.getElementById("cfgStreamProxyUrl"),
      cfgStreamProxyToken: document.getElementById("cfgStreamProxyToken"),
      cfgCorsProxyEnabled: document.getElementById("cfgCorsProxyEnabled"),
      cfgCorsProxyUrl: document.getElementById("cfgCorsProxyUrl"),
      cfgDirectEnabled: document.getElementById("cfgDirectEnabled"),
      cfgDirectBaseURL: document.getElementById("cfgDirectBaseURL"),
      cfgDirectApiKey: document.getElementById("cfgDirectApiKey"),
      agentBanner: document.getElementById("agentBanner"),
      agentBannerText: document.getElementById("agentBannerText"),
      chatContainer: document.getElementById("chatContainer"),
      emptyHint: document.getElementById("emptyHint"),
      btnClearChat: document.getElementById("btnClearChat"),
      pendingCommentsPanel: document.getElementById("pendingCommentsPanel"),
      pendingCommentsList: document.getElementById("pendingCommentsList"),
      pendingCommentsCount: document.getElementById("pendingCommentsCount"),
      slashMenu: document.getElementById("slashMenu"),
      slashMenuList: document.getElementById("slashMenuList"),
      chatInput: document.getElementById("chatInput"),
      btnSend: document.getElementById("btnSend"),
      // Text Analytics Modal
      // Settings Modal
      settingsModal: document.getElementById("settingsModal"),
      btnCloseSettings: document.getElementById("btnCloseSettings"),
      btnSaveSettings: document.getElementById("btnSaveSettings"),
      cfgHouseVoice: document.getElementById("cfgHouseVoice"),
      licEmailDisplay: document.getElementById("licEmailDisplay"),
      licDeviceDisplay: document.getElementById("licDeviceDisplay"),
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
    const nextTheme = this.state.theme === "dark" ? "light" : "dark";
    this.setTheme(nextTheme);
  }

  bindEvents() {
    // Theme toggle
    this.dom.btnToggleTheme.addEventListener("click", () => this.toggleTheme());

    // Composer events
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

    // Auto-expand textarea & check for slash command
    this.dom.chatInput.addEventListener("input", () => {
      this.dom.chatInput.style.height = "auto";
      this.dom.chatInput.style.height = Math.min(this.dom.chatInput.scrollHeight, 120) + "px";

      const val = this.dom.chatInput.value;
      if (val.startsWith("/")) {
        this.openSlashMenu(val.slice(1).toLowerCase());
      } else {
        this.closeSlashMenu();
      }
    });

    // Close slash menu when clicking outside
    document.addEventListener("click", (e) => {
      if (!e.target.closest(".input-box-wrapper")) {
        this.closeSlashMenu();
      }
    });

    // Toolbar buttons
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
    if (this.dom.settingsModal) {
      this.dom.settingsModal.addEventListener("click", (e) => {
        if (e.target === this.dom.settingsModal) this.closeSettings();
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
    this.dom.btnSaveSettings.addEventListener("click", () => this.saveSettings());
    this.dom.settingsModal.addEventListener("click", (e) => {
      if (e.target === this.dom.settingsModal) this.closeSettings();
    });

    // Text Analytics Modal

    // License Modal
    if (this.dom.btnActivate) {
      this.dom.btnActivate.addEventListener("click", () => this.handleActivate());
    }
    if (this.dom.licKey) {
      this.dom.licKey.addEventListener("keydown", (e) => {
        if (e.key === "Enter") this.handleActivate();
      });
    }
    if (this.dom.licEmail) {
      this.dom.licEmail.addEventListener("keydown", (e) => {
        if (e.key === "Enter") this.dom.licKey?.focus();
      });
    }
  }

  /**
   * Tải danh sách skill từ skills.json (static). Lỗi → fallback DEFAULT_SKILLS.
   */
  async loadSkills() {
    try {
      const res = await fetch("skills.json");
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

  renderSkills() {
    // Skills bar đã bỏ khỏi UI — slash menu (/) dùng data từ state.skills
    // Method này giữ để tương thích nhưng là no-op.
  }

  initOffice() {
    if (typeof Office !== "undefined" && Office.onReady) {
      Office.onReady((info) => {
        if (info && info.host === Office.HostType.Word) {
          this.updateStatus("online", "Word Ready");
        } else {
          this.updateStatus("busy", "Web Preview");
        }
      }).catch((err) => {
        console.warn("[Ori AI] Office.onReady catch:", err);
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

  // ─── LICENSE LIFECYCLE ────────────────────────────────────────────────

  /**
   * Kiểm tra license local; nếu chưa có thì show modal và đợi user activate.
   * Trả về true nếu đã activated (sẵn sàng dùng app).
   */
  async initLicense() {
    // Đảm bảo deviceId đã được tạo (UUID v4)
    getOrCreateDeviceId();

    // Fetch public config để lấy default effort từ server
    const publicConfig = await licenseClient.fetchPublicConfig();
    if (publicConfig?.aiDefaultEffort) {
      // Chỉ set nếu user chưa có lựa chọn
      const saved = docState.loadEffort(null);
      if (!saved) {
        this.state.config.effort = publicConfig.aiDefaultEffort;
        docState.saveEffort(publicConfig.aiDefaultEffort);
      }
    }

    // Nếu server không yêu cầu license (dev mode), coi như đã activate
    if (publicConfig && publicConfig.licenseRequired === false) {
      return true;
    }

    // Check local license
    if (licenseClient.isActivated()) {
      // Bắt đầu heartbeat
      licenseClient.startHeartbeat();
      licenseClient.onInvalidate(() => this.showLicenseModal(true));
      return true;
    }

    // Chưa có → show modal
    this.showLicenseModal(false);
    return false;
  }

  showLicenseModal(showStatus) {
    if (!this.dom.licenseModal) return;
    this.dom.licenseModal.classList.add("open");
    this.dom.licenseModal.style.display = "flex";
    if (showStatus && this.dom.licStatus) {
      this.showLicenseStatus("error", "Phiên bản quyền đã hết hạn. Vui lòng kích hoạt lại.");
    }
  }

  hideLicenseModal() {
    if (!this.dom.licenseModal) return;
    this.dom.licenseModal.classList.remove("open");
    this.dom.licenseModal.style.display = "none";
  }

  showLicenseStatus(type, msg) {
    if (!this.dom.licStatus) return;
    this.dom.licStatus.className = `license-status show ${type}`;
    this.dom.licStatus.textContent = msg;
  }

  clearLicenseStatus() {
    if (!this.dom.licStatus) return;
    this.dom.licStatus.className = "license-status";
    this.dom.licStatus.textContent = "";
  }

  async handleActivate() {
    const email = this.dom.licEmail.value.trim();
    const licenseKey = this.dom.licKey.value.trim();

    // Validate format
    if (!email || !licenseKey) {
      this.showLicenseStatus("error", "Vui lòng nhập đầy đủ email và License Key.");
      return;
    }
    const uuidRegex = /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/;
    if (!uuidRegex.test(licenseKey)) {
      this.showLicenseStatus("error", "License Key phải là UUID (8-4-4-4-12 hex chars).");
      return;
    }

    this.dom.btnActivate.disabled = true;
    this.dom.btnActivate.textContent = "Đang kích hoạt...";
    this.clearLicenseStatus();

    try {
      const deviceName = navigator.userAgent.includes("Word") ? "Word Desktop" : "Web";
      const result = await licenseClient.activate({ email, licenseKey, deviceName });

      this.showLicenseStatus("success", `✅ Kích hoạt thành công! ${result.isNewDevice ? "Đã đăng ký thiết bị mới." : "Thiết bị đã được đăng ký trước đó."}`);

      // Bắt đầu heartbeat
      licenseClient.startHeartbeat();
      licenseClient.onInvalidate(() => this.showLicenseModal(true));

      // Reload app: ẩn modal, init các phần còn lại
      setTimeout(() => {
        this.hideLicenseModal();
        this.initWebSocket();
        this.loadState();
        this.renderSkills();
        this.initEffortDropdown();
        this.showToast("✅ Kích hoạt bản quyền thành công!");
      }, 800);
    } catch (err) {
      this.showLicenseStatus("error", err.message || "Kích hoạt thất bại. Vui lòng thử lại.");
    } finally {
      this.dom.btnActivate.disabled = false;
      this.dom.btnActivate.textContent = "Kích hoạt";
    }
  }

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

  /**
   * Lưu session hiện tại vào IndexedDB (word-GPT-Plus checkpointer pattern).
   * Mỗi file Word = 1 thread. Khi mở lại file → load lại session.
   * Fallback to localStorage nếu IndexedDB bị chặn.
   */
  async _persistSession() {
    try {
      const docPath = await wordBridge.getDocumentPathSafe?.() || "unknown";
      const systemPrompt = docState.loadSystemPrompt();
      await saveSession({
        threadId: this._threadId,
        docPath,
        messages: this.state.messages,
        systemPrompt,
        config: this.state.config,
      });
    } catch (err) {
      console.warn("[Ori AI] persistSession failed:", err.message);
    }
  }

  /**
   * Lấy session gần nhất của file Word hiện tại (nếu có).
   */
  async _loadSessionForCurrentDoc() {
    try {
      const docPath = await wordBridge.getDocumentPathSafe?.() || "unknown";
      const session = await loadLatestSession(docPath);
      if (!session) return false;
      this._threadId = session.threadId;
      this.state.messages = session.messages || [];
      this.state.config = { ...this.state.config, ...(session.config || {}) };
      if (this.state.messages.length > 0) {
        this.dom.emptyHint.style.display = "none";
        for (const msg of this.state.messages) {
          this.renderMessageBubble(msg.role, msg.displayContent || msg.content, {
            thinking: msg.thinking,
            persist: false,
          });
        }
        this.showToast(`📂 Đã khôi phục ${this.state.messages.length} tin nhắn từ phiên trước.`);
      }
      return true;
    } catch (err) {
      console.warn("[Ori AI] loadSessionForCurrentDoc failed:", err.message);
      return false;
    }
  }

  /**
   * Debounced persist session (1s) — tránh ghi liên tục khi đang chat.
   */
  _schedulePersistSession() {
    if (this._saveSessionDebounced) clearTimeout(this._saveSessionDebounced);
    this._saveSessionDebounced = setTimeout(() => this._persistSession(), 1000);
  }

  // ─── CONTEXT (AI Agent luôn tự đọc Toàn bài) ──────────────────────

  /**
   * AI Agent luôn tự động đọc toàn bộ tài liệu Word đang mở làm context nền.
   * User KHÔNG cần bấm "Ghim" gì cả.
   */
  async assembleAutoContext() {
    let fullDocText = "";
    try {
      fullDocText = await wordBridge.getFullDocumentText(30000);
    } catch (_) {}
    return {
      fullDocText: fullDocText ? fullDocText.trim() : "",
    };
  }

  // ─── SKILLS TRIGGER ───

  async triggerSkill(skill) {
    const { fullDocText } = await this.assembleAutoContext();

    let promptText = skill.prompt;
    if (fullDocText) {
      promptText = `${skill.prompt}\n\n[Tài liệu Word đang mở (${wordBridge.getTextAnalytics(fullDocText).words} từ)]:\n"""\n${fullDocText}\n"""`;
    }

    this.sendMessage(promptText, {
      displayPrompt: `${skill.icon} ${skill.label}`,
      skillId: skill.id,
      defaultMode: skill.defaultMode,
      originalText: fullDocText || "",
    });
  }

  // ─── CHAT & STREAMING (luôn đính kèm Toàn bài) ───

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

    // AI Agent luôn tự đọc toàn bộ tài liệu
    const { fullDocText } = await this.assembleAutoContext();

    let fullPrompt = userInput;
    if (fullDocText) {
      fullPrompt = `${userInput}\n\n[Tài liệu Word đang mở (${wordBridge.getTextAnalytics(fullDocText).words} từ)]:\n"""\n${fullDocText}\n"""`;
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

    // 2. Assistant Bubble + Agent Banner
    const assistantBubble = this.createAssistantBubble();
    this.setStreamingState(true);
    this.setAgentActive(true, meta?.skillId ? `đang thực hiện kỹ năng` : `đang phân tích tài liệu`);

    // 3. System Prompt & House Voice
    let systemPrompt = docState.loadSystemPrompt();
    const houseVoice = this.state.config.houseVoice || docState.loadHouseVoice();
    if (houseVoice && houseVoice.trim()) {
      systemPrompt += `\n\n<personal_writing_style>\n${houseVoice.trim()}\nQuy tắc: Luôn mô phỏng đúng văn phong, nhịp điệu và cách dùng từ của đoạn mẫu trên.\n</personal_writing_style>`;
    }

    // 4. Tracking marker đã spawn để tránh spawn trùng
    const spawnedMarkers = new Set();
    const pendingCommentsThisTurn = [];

    const handleNewMarkers = async (fullText) => {
      const markers = this.extractCommentMarkers(fullText);
      for (const marker of markers) {
        const key = `${marker.anchor}|||${marker.suggestion}`;
        if (spawnedMarkers.has(key)) continue;
        spawnedMarkers.add(key);
        // Spawn comment lên Word
        try {
          await wordBridge.insertComment(
            `Góp ý: ${marker.anchor}\n→ Sửa thành: ${marker.suggestion}`,
            marker.anchor,
            { authorTag: true }
          );
          pendingCommentsThisTurn.push(marker);
          this.addPendingComment(marker);
        } catch (err) {
          console.warn("[Ori AI] spawn comment thất bại:", err.message);
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
          // Real-time spawn comment khi AI generate marker
          handleNewMarkers(fullText);
          this.scrollToBottom();
        },

        onDone: (fullText, fullThinking, aborted) => {
          this.setStreamingState(false);
          this.setAgentActive(false);
          // Đảm bảo scan lần cuối cho marker bị stream cắt
          handleNewMarkers(fullText);
          assistantBubble.finalize(fullText, fullThinking, meta);

          this.state.messages.push({
            role: "assistant",
            content: fullText,
            thinking: fullThinking,
          });
          docState.saveHistory(this.state.messages);
          this._schedulePersistSession();
          this.scrollToBottom();

          // Nếu có comment pending, thông báo
          if (pendingCommentsThisTurn.length > 0) {
            this.showToast(
              `💬 ${pendingCommentsThisTurn.length} nhận xét đã ghim vào lề Word. Xem panel bên dưới để duyệt.`
            );
          }
        },

        onError: (err) => {
          this.setStreamingState(false);
          this.setAgentActive(false);
          assistantBubble.showError(err.message);
          this.scrollToBottom();
        },
      });
    } catch (_) {
      this.setStreamingState(false);
      this.setAgentActive(false);
    }
  }

  // ─── COMMENT MARKER PARSING (real-time) ───────────────────────────────

  /**
   * Tách tất cả marker [Góp ý: "anchor" -> suggestion] từ text.
   * Khác extractComments() ở chỗ: chỉ lấy anchor + suggestion (không gộp replies).
   */
  extractCommentMarkers(text) {
    if (!text || typeof text !== "string") return [];
    const out = [];
    const regex = /\[(?:Góp ý|Comment|Nhận xét|Feedback):\s*"([^"]{3,300})"\s*(?:->|:|➔)\s*([^\]]+)\]/gim;
    let m;
    while ((m = regex.exec(text)) !== null) {
      out.push({
        anchor: m[1].trim(),
        suggestion: m[2].trim(),
      });
    }
    return out;
  }

  // ─── PENDING COMMENTS PANEL ───────────────────────────────────────────

  addPendingComment(marker) {
    if (!this.dom.pendingCommentsPanel) return;
    if (this.dom.pendingCommentsPanel.style.display === "none") {
      this.dom.pendingCommentsPanel.style.display = "block";
    }
    if (this.dom.pendingCommentsList) {
      const id = `pending_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
      const card = document.createElement("div");
      card.className = "pending-comment-card";
      card.dataset.id = id;
      card.dataset.anchor = marker.anchor;
      card.dataset.suggestion = marker.suggestion;
      card.innerHTML = `
        <div class="pending-comment-anchor">📌 "${this.escapeHtml(marker.anchor.slice(0, 50))}${marker.anchor.length > 50 ? "..." : ""}"</div>
        <div class="pending-comment-suggestion">→ ${this.escapeHtml(marker.suggestion.slice(0, 80))}${marker.suggestion.length > 80 ? "..." : ""}</div>
        <div class="pending-comment-actions">
          <button class="btn-accept" data-action="accept">✓ Áp dụng</button>
          <button class="btn-reject" data-action="reject">✕ Bỏ qua</button>
          <button class="btn-reply" data-action="reply">💬 Reply</button>
        </div>
      `;
      card.querySelector('[data-action="accept"]').addEventListener("click", () => this.handleAcceptComment(card));
      card.querySelector('[data-action="reject"]').addEventListener("click", () => this.handleRejectComment(card));
      card.querySelector('[data-action="reply"]').addEventListener("click", () => this.handleReplyComment(card));
      this.dom.pendingCommentsList.appendChild(card);
      this.updatePendingCount();
    }
  }

  updatePendingCount() {
    if (!this.dom.pendingCommentsCount) return;
    const n = this.dom.pendingCommentsList ? this.dom.pendingCommentsList.children.length : 0;
    this.dom.pendingCommentsCount.textContent = n;
  }

  async handleAcceptComment(card) {
    const anchor = card.dataset.anchor;
    const suggestion = card.dataset.suggestion;
    card.querySelectorAll("button").forEach((b) => (b.disabled = true));
    try {
      await wordBridge.acceptComment({ anchorText: anchor, replacementText: suggestion });
      this.showToast("✅ Đã áp dụng sửa đổi & xoá nhận xét.");
      card.remove();
      this.updatePendingCount();
    } catch (err) {
      card.querySelectorAll("button").forEach((b) => (b.disabled = false));
      this.showToast(`⚠️ Lỗi: ${err.message}`);
    }
  }

  async handleRejectComment(card) {
    const anchor = card.dataset.anchor;
    card.querySelectorAll("button").forEach((b) => (b.disabled = true));
    try {
      await wordBridge.deleteComment({ anchorText: anchor });
      this.showToast("🗑️ Đã bỏ qua & xoá nhận xét.");
      card.remove();
      this.updatePendingCount();
    } catch (err) {
      // Vẫn xoá khỏi panel dù Word có lỗi
      card.remove();
      this.updatePendingCount();
    }
  }

  async handleReplyComment(card) {
    const anchor = card.dataset.anchor;
    const reply = window.prompt(`Reply cho nhận xét về:\n"${anchor.slice(0, 60)}..."\n\n(Nhập nội dung reply)`, "");
    if (!reply || !reply.trim()) return;
    try {
      await wordBridge.replyToComment({ anchorText: anchor, replyText: `[User] ${reply.trim()}` });
      this.showToast("💬 Đã reply vào comment.");
    } catch (err) {
      this.showToast(`⚠️ Lỗi reply: ${err.message}`);
    }
  }

  setAgentActive(active, toolName = "") {
    if (!this.dom.agentBanner) return;
    if (active) {
      this.dom.agentBanner.classList.add("show");
      this.dom.agentBannerText.textContent = `🤖 Ori AI Agent ${toolName}...`;
    } else {
      this.dom.agentBanner.classList.remove("show");
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

  renderMessageBubble(role, content, { thinking = "", persist = true, meta = {} } = {}) {
    const bubble = document.createElement("div");
    bubble.className = `message-bubble ${role}`;

    if (role === "user") {
      // Strip leading emoji + khoảng trắng để bubble user gọn hơn (icon đã có ở starter card)
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

      // Thêm action buttons
      this.attachActionButtons(bubble, content, meta);
      this.attachAnchorChips(bubble, content);

      // Tự động phát hiện và hiển thị Widget Bình luận lề trang
      const comments = this.extractComments(content);
      if (comments.length > 0) {
        this.attachCommentsWidget(bubble, comments);
      }

      // Tự động hiển thị Word-by-Word Diff nếu có ngữ cảnh so sánh
      if (meta?.originalText) {
        const cleanText = this.extractExecutableText(content);
        this.attachDiffPreview(bubble, meta.originalText, cleanText);
      }
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

      finalize: (fullText, fullThinking, meta = {}) => {
        contentDiv.innerHTML = this.formatMarkdown(fullText);
        this.attachActionButtons(bubble, fullText, meta);
        this.attachAnchorChips(bubble, fullText);

        // Tự động phát hiện và render Widget Góp ý lề trang
        const comments = this.extractComments(fullText);
        if (comments.length > 0) {
          this.attachCommentsWidget(bubble, comments);
        }

        // Tự động tính toán và hiển thị Word-Level Diffing
        if (meta?.originalText) {
          const cleanText = this.extractExecutableText(fullText);
          this.attachDiffPreview(bubble, meta.originalText, cleanText);
        }
      },

      showError: (errText) => {
        contentDiv.innerHTML += `<div style="color: var(--danger); margin-top: 6px; font-weight: 500;">⚠️ ${this.escapeHtml(errText)}</div>`;
      },
    };
  }

  /**
   * Trích xuất các bình luận có cấu trúc từ phản hồi của AI
   */
  extractComments(text) {
    if (!text || typeof text !== "string") return [];
    const comments = [];

    // Cú pháp 1: [Góp ý: "câu trích" -> nội dung]
    const regex1 = /\[(?:Góp ý|Comment|Nhận xét|Feedback):\s*"([^"]{4,300})"\s*(?:->|:|➔)\s*([^\]]+)\]/gim;
    let match;
    while ((match = regex1.exec(text)) !== null) {
      comments.push({
        anchor: match[1].trim(),
        text: match[2].trim(),
      });
    }

    // Cú pháp 2: [[comment: anchor="..." | text="..."]]
    const regex2 = /\[\[comment:\s*anchor="([^"]+)"\s*\|\s*text="([^"]+)"]]/gim;
    while ((match = regex2.exec(text)) !== null) {
      comments.push({
        anchor: match[1].trim(),
        text: match[2].trim(),
      });
    }

    return comments;
  }

  // attachCommentsWidget cũ — bỏ vì streaming comment spawn trong sendMessage
  // đã tự động tạo comment ở Word + đẩy vào pending comments panel.
  // Method này giữ để tương thích, là no-op.
  attachCommentsWidget(container, comments) {
    /* streaming đã lo — no-op */
  }

  attachDiffPreview(container, originalText, cleanText) {
    if (!originalText || !cleanText || originalText.trim() === cleanText.trim()) return;

    // Diff preview dạng gọn — không còn nút "Vạch đỏ", chỉ hiển thị diff để user xem.
    const diffChunks = computeWordDiff(originalText, cleanText);
    const changeCount = diffChunks.filter((c) => c.type !== "equal").length;
    if (changeCount === 0) return;

    const diffBox = document.createElement("details");
    diffBox.className = "diff-box";
    diffBox.innerHTML = `
      <summary>
        <span>👁️ So sánh sửa đổi (${changeCount} điểm)</span>
      </summary>
      <div class="diff-body">${renderDiffHtml(diffChunks)}</div>
    `;

    const actionsBar = container.querySelector(".msg-actions");
    container.insertBefore(diffBox, actionsBar);
  }

  /**
   * Trích xuất đoạn văn bản thuần túy cần thao tác (bỏ qua phần giải thích chào hỏi)
   */
  extractExecutableText(text) {
    if (!text || typeof text !== "string") return "";

    // 1. Nếu có phần văn bản nằm giữa hai dòng phân cách --- ... ---
    const delimiterMatch = text.match(/(?:^|\n)\s*---\s*\n([\s\S]*?)\n\s*---\s*(?:\n|$)/);
    if (delimiterMatch && delimiterMatch[1] && delimiterMatch[1].trim()) {
      return delimiterMatch[1].trim();
    }

    // 2. Nếu có code block ```...```
    const codeBlockMatch = text.match(/```(?:[a-z]*)\n([\s\S]*?)```/i);
    if (codeBlockMatch && codeBlockMatch[1] && codeBlockMatch[1].trim()) {
      return codeBlockMatch[1].trim();
    }

    // 3. Fallback: dùng toàn bộ text
    return text.trim();
  }

  // ─── ACTION BUTTONS (chỉ giữ Copy) ─────────────────────────────────

  attachActionButtons(container, text, meta = {}) {
    const actionsBar = document.createElement("div");
    actionsBar.className = "msg-actions";
    const cleanText = this.extractExecutableText(text);

    const actions = [
      {
        id: "copy",
        label: "📋 Sao chép",
        fn: () => {
          navigator.clipboard.writeText(cleanText);
          this.showToast("📋 Đã sao chép vào bộ nhớ tạm!");
        },
      },
    ];

    for (const act of actions) {
      const btn = document.createElement("button");
      btn.className = "action-btn";
      btn.textContent = act.label;
      btn.addEventListener("click", async () => {
        try {
          await act.fn();
        } catch (err) {
          console.error(`[Ori AI] action "${act.id}" failed:`, err);
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
            // Visual feedback (không cần toast — user thích tính năng này không cần thông báo)
            chip.classList.add("jump-success");
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

  // quoteText cũ — bỏ vì skills bar và context chips đã bỏ.
  quoteText(text) { /* no-op */ }

  // ─── SLASH COMMANDS MENU ───

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
    if (!Array.isArray(this._slashFilteredSkills)) this._slashFilteredSkills = [];
    this.dom.slashMenuList.innerHTML = "";

    this._slashFilteredSkills.forEach((skill, idx) => {
      const item = document.createElement("div");
      item.className = `slash-item ${idx === this._slashSelectedIndex ? "selected" : ""}`;
      item.innerHTML = `
        <span class="slash-item-icon">${skill.icon}</span>
        <span class="slash-item-label">${this.escapeHtml(skill.label)}</span>
        <span class="slash-item-tip">${this.escapeHtml(skill.tip)}</span>
      `;
      item.addEventListener("click", () => {
        this.executeSlashSkill(skill);
      });
      this.dom.slashMenuList.appendChild(item);
    });
  }

  navigateSlashMenu(direction) {
    if (!this._slashFilteredSkills || this._slashFilteredSkills.length === 0) return;
    this._slashSelectedIndex = (this._slashSelectedIndex + direction + this._slashFilteredSkills.length) % this._slashFilteredSkills.length;
    this.renderSlashMenuItems();
  }

  selectCurrentSlashItem() {
    if (this._slashFilteredSkills && this._slashFilteredSkills[this._slashSelectedIndex]) {
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

  // ─── SETTINGS MODAL & TEST CONNECTION ───

  openSettings() {
    this.dom.cfgHouseVoice.value = this.state.config.houseVoice || "";

    // Load 3 layer config (stream-factory)
    const all = getAllStreamConfigs();
    this.dom.cfgStreamProxyEnabled.checked = all.streamProxy.enabled || false;
    this.dom.cfgStreamProxyUrl.value = all.streamProxy.url || "";
    this.dom.cfgStreamProxyToken.value = all.streamProxy.token || "";
    this.dom.cfgCorsProxyEnabled.checked = all.corsProxy.enabled || false;
    this.dom.cfgCorsProxyUrl.value = all.corsProxy.url || "";
    this.dom.cfgDirectEnabled.checked = !!all.providerKey.apiKey;
    this.dom.cfgDirectBaseURL.value = all.providerKey.baseURL || "";
    this.dom.cfgDirectApiKey.value = all.providerKey.apiKey || "";

    // Hiển thị thông tin license (sẽ được update khi license client ready)
    const lic = docState.loadLicense();
    if (this.dom.licEmailDisplay) {
      this.dom.licEmailDisplay.textContent = lic?.email || "—";
    }
    if (this.dom.licDeviceDisplay) {
      this.dom.licDeviceDisplay.textContent = lic?.deviceId
        ? `${lic.deviceId.substring(0, 12)}...`
        : "—";
    }

    this.dom.settingsModal.classList.add("open");
  }

  closeSettings() {
    this.dom.settingsModal.classList.remove("open");
  }

  saveSettings() {
    this.state.config.houseVoice = this.dom.cfgHouseVoice.value.trim();
    docState.saveHouseVoice(this.state.config.houseVoice);

    // Save 3 layer config (stream-factory)
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

  // ─── SESSIONS DIALOG (IndexedDB checkpointer) ────────────────────────

  async openSessions() {
    this.dom.sessionsModal.classList.add("open");
    await this._loadSessionsList();
  }

  closeSessions() {
    this.dom.sessionsModal.classList.remove("open");
  }

  async _loadSessionsList() {
    if (!this.dom.sessionsList) return;
    this.dom.sessionsList.innerHTML = '<div class="sessions-empty">Đang tải...</div>';

    const sessions = await listSessions(50);
    if (!sessions.length) {
      this.dom.sessionsList.innerHTML = '<div class="sessions-empty">Chưa có phiên nào được lưu.</div>';
      return;
    }

    this.dom.sessionsList.innerHTML = "";
    for (const s of sessions) {
      const item = document.createElement("div");
      item.className = "session-item";
      const dt = new Date(s.updatedAt);
      const timeStr = dt.toLocaleString("vi-VN");
      const msgCount = (s.messages || []).length;
      item.innerHTML = `
        <div class="session-info">
          <div class="session-title">${this.escapeHtml(s.title || s.docPath || "Phiên")}</div>
          <div class="session-meta">${this.escapeHtml(s.docPath || "")} · ${msgCount} tin nhắn · ${timeStr}</div>
        </div>
        <button class="session-delete" title="Xoá phiên này">🗑️</button>
      `;
      item.addEventListener("click", async (e) => {
        if (e.target.classList.contains("session-delete")) return;
        await this._restoreSession(s);
      });
      item.querySelector(".session-delete").addEventListener("click", async (e) => {
        e.stopPropagation();
        if (confirm(`Xoá phiên "${s.title || s.docPath}"?`)) {
          await deleteSession(s.threadId);
          await this._loadSessionsList();
          this.showToast("🗑️ Đã xoá phiên.");
        }
      });
      this.dom.sessionsList.appendChild(item);
    }
  }

  async _restoreSession(session) {
    this._threadId = session.threadId;
    this.state.messages = session.messages || [];
    if (session.config) this.state.config = { ...this.state.config, ...session.config };
    // Re-render chat
    this.dom.chatContainer.innerHTML = "";
    if (this.state.messages.length === 0) {
      // restore empty hint
      this.dom.emptyHint.style.display = "flex";
    } else {
      this.dom.emptyHint.style.display = "none";
      for (const msg of this.state.messages) {
        this.renderMessageBubble(msg.role, msg.displayContent || msg.content, {
          thinking: msg.thinking,
          persist: false,
        });
      }
    }
    this.closeSessions();
    this.showToast(`📂 Đã khôi phục ${this.state.messages.length} tin nhắn.`);
  }

  clearChatHistory() {
    if (this._clearConfirmTimer) {
      clearTimeout(this._clearConfirmTimer);
      this._clearConfirmTimer = null;
      this.state.messages = [];
      docState.clearHistory();
      this.dom.chatContainer.innerHTML = "";
      this.dom.chatContainer.appendChild(this.dom.emptyHint);
      this.dom.emptyHint.style.display = "flex";
      this.dom.btnClearChat.innerHTML = '<span class="tool-icon">⌫</span> Xóa chat';
      this.showToast("🗑️ Đã xóa lịch sử trò chuyện.");
      return;
    }

    this.dom.btnClearChat.innerHTML = '<span class="tool-icon">⚠️</span> Xác nhận?';
    this.showToast("Nhấp lại 'Xác nhận?' để xóa lịch sử chat.");

    this._clearConfirmTimer = setTimeout(() => {
      this.dom.btnClearChat.innerHTML = '<span class="tool-icon">⌫</span> Xóa chat';
      this._clearConfirmTimer = null;
    }, 3000);
  }

  // ─── RICH MARKDOWN PARSER ───

  formatMarkdown(text) {
    if (!text) return "";
    // Tokenize: escape toàn bộ text trước, sau đó thay thế các pattern
    // bằng placeholder đặc biệt để tránh escape 2 lần.
    const escaped = this.escapeHtml(text);

    // Dùng 1 pass duy nhất qua escaped text. Block-level (h1, pre) cần
    // bảo toàn nội dung bên trong, nên chỉ áp dụng <br/> bên ngoài block.
    let html = escaped;

    // Code blocks (must come FIRST để bảo vệ nội dung bên trong)
    html = html.replace(/```([a-z]*)\n([\s\S]*?)```/gim, (m, lang, body) => {
      return `<pre><code>${body}</code></pre>`;
    });

    // Headers
    html = html.replace(/^### (.*$)/gim, "<h3>$1</h3>");
    html = html.replace(/^## (.*$)/gim, "<h2>$1</h2>");
    html = html.replace(/^# (.*$)/gim, "<h1>$1</h1>");

    // Inline code
    html = html.replace(/`([^`]+)`/g, "<code>$1</code>");

    // Bold, Italic (bold trước vì ** chứa *)
    html = html.replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>");
    html = html.replace(/\*(.*?)\*/g, "<em>$1</em>");

    // Lists: gom các dòng liên tiếp bắt đầu bằng `-` hoặc `*` thành <ul>
    html = html.replace(/(^|\n)((?:[-*]\s+[^\n]+\n?)+)/g, (match, prefix, block) => {
      const items = block.trim().split(/\n/).map((l) => `<li>${l.replace(/^[-*]\s+/, "")}</li>`).join("");
      return `${prefix}<ul>${items}</ul>`;
    });

    // Line breaks: chỉ thay \n còn lại (không nằm trong <pre>...</pre>)
    // Tách ra theo dòng rồi nối lại, bỏ qua nội dung trong <pre>
    const parts = html.split(/(<pre>[\s\S]*?<\/pre>)/g);
    html = parts
      .map((part) => {
        if (part.startsWith("<pre>")) return part;
        return part.replace(/\n/g, "<br/>");
      })
      .join("");

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

// Khởi chạy ứng dụng
const app = new AppController();

if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady().catch((e) => console.warn("[Ori AI] Top-level Office.onReady error:", e));
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", () => app.init());
} else {
  app.init();
}
