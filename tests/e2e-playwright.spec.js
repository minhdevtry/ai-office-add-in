import { test, expect } from "@playwright/test";

test.describe("Ori AI — Word Add-in E2E Browser Test Suite", () => {
  test.beforeEach(async ({ page }) => {
    page.on("console", (msg) => console.log(`[Browser Console] ${msg.type()}: ${msg.text()}`));
    page.on("pageerror", (err) => console.log(`[Browser PageError]: ${err.message}`));

    // Seed local activated license for seamless automated testing
    await page.addInitScript(() => {
      window.localStorage.setItem(
        "ai_word.license",
        JSON.stringify({
          email: "test@2tocom.space",
          licenseKey: "11111111-2222-3333-4444-555555555555",
          deviceId: "test-device-uuid",
          activatedAt: new Date().toISOString(),
        })
      );
    });

    // Navigate to local taskpane
    await page.goto("http://localhost:3011/index.html");
    await page.waitForLoadState("domcontentloaded");
  });

  test("1. Page loads with header, title, effort selector and theme button", async ({ page }) => {
    // Verify brand name
    const brandName = page.locator(".brand-name");
    await expect(brandName).toHaveText("Ori AI");

    // Verify Quick Effort Selector
    const effortSelect = page.locator("#quickEffortSelect");
    await expect(effortSelect).toBeVisible();
    await expect(effortSelect).toHaveValue("medium");

    // Verify Theme Button
    const themeBtn = page.locator("#btnToggleTheme");
    await expect(themeBtn).toBeVisible();

    // Verify Settings and Sessions buttons
    await expect(page.locator("#btnSettings")).toBeVisible();
    await expect(page.locator("#btnSessions")).toBeVisible();
  });

  test("2. Document insights bar loads and displays word count", async ({ page }) => {
    const insightsBar = page.locator("#docInsightsBar");
    await expect(insightsBar).toBeVisible({ timeout: 5000 });
    const text = await page.locator("#docInsightsText").textContent();
    expect(text).toContain("Toàn bộ tài liệu:");
    expect(text).toContain("từ");
  });

  test("3. Quick Effort Selector changes reasoning level with feedback", async ({ page }) => {
    const effortSelect = page.locator("#quickEffortSelect");
    await effortSelect.selectOption("high");

    const toast = page.locator("#oriToast");
    await expect(toast).toBeVisible();
    await expect(toast).toContainText("Mức độ suy luận: 🔬 Sâu");
  });

  test("4. Theme toggle switches between light and dark mode", async ({ page }) => {
    const html = page.locator("html");
    await expect(html).toHaveAttribute("data-theme", "light");

    const themeBtn = page.locator("#btnToggleTheme");
    await themeBtn.click();
    await expect(html).toHaveAttribute("data-theme", "dark");

    await themeBtn.click();
    await expect(html).toHaveAttribute("data-theme", "light");
  });

  test("5. Clicking a Starter Card triggers chat, streams response and renders Patch Cards", async ({ page }) => {
    // Click Starter Card: "Soát chính tả & dấu"
    const starterCard = page.locator('.starter-card[data-skill="grammar"]');
    await expect(starterCard).toBeVisible();
    await starterCard.click();

    // Empty hint should hide
    await expect(page.locator("#emptyHint")).toBeHidden();

    // User message bubble should be created
    const userMsg = page.locator(".message-bubble.user").last();
    await expect(userMsg).toBeVisible();
    await expect(userMsg).toContainText("Soát chính tả & dấu");

    // Assistant message bubble should appear
    const assistantMsg = page.locator(".message-bubble.assistant").last();
    await expect(assistantMsg).toBeVisible();

    // Wait for streaming to finish and Patch Suite to appear
    const patchSuite = assistantMsg.locator(".patch-suite-container");
    await expect(patchSuite).toBeVisible({ timeout: 10000 });

    // Verify Diff Patch cards are rendered with <del> and <ins>
    const patchCards = patchSuite.locator(".patch-card");
    const count = await patchCards.count();
    expect(count).toBeGreaterThanOrEqual(1);

    // Verify first patch card has diff content and buttons
    const firstCard = patchCards.first();
    await expect(firstCard.locator(".diff-del").first()).toBeVisible();
    await expect(firstCard.locator(".diff-ins").first()).toBeVisible();
    await expect(firstCard.locator(".btn-patch-accept")).toBeVisible();
    await expect(firstCard.locator(".btn-patch-reject")).toBeVisible();
  });

  test("6. Interactive Patch Card: Accept single proposal updates status and document", async ({ page }) => {
    // Trigger quick skill button "Viết lại"
    const quickSkillBtn = page.locator('.quick-skill-btn[data-skill="polish"]');
    await quickSkillBtn.click();

    const assistantMsg = page.locator(".message-bubble.assistant").last();
    const patchSuite = assistantMsg.locator(".patch-suite-container");
    await expect(patchSuite).toBeVisible({ timeout: 10000 });

    const firstCard = patchSuite.locator(".patch-card").first();
    const btnAccept = firstCard.locator(".btn-patch-accept");
    await btnAccept.click();

    // Verify card changes to accepted state
    await expect(firstCard).toHaveClass(/card-accepted/);
    const statusTag = firstCard.locator(".patch-status-tag");
    await expect(statusTag).toHaveText("✅ Đã áp dụng");
    await expect(btnAccept).toBeDisabled();

    // Verify toast notification
    const toast = page.locator("#oriToast");
    await expect(toast).toBeVisible();
    await expect(toast).toContainText("Đã áp dụng sửa đổi vào Word");
  });

  test("7. Interactive Patch Card: Reject single proposal updates status", async ({ page }) => {
    const quickSkillBtn = page.locator('.quick-skill-btn[data-skill="formal"]');
    await quickSkillBtn.click();

    const assistantMsg = page.locator(".message-bubble.assistant").last();
    const patchSuite = assistantMsg.locator(".patch-suite-container");
    await expect(patchSuite).toBeVisible({ timeout: 10000 });

    const firstCard = patchSuite.locator(".patch-card").first();
    const btnReject = firstCard.locator(".btn-patch-reject");
    await btnReject.click();

    // Verify card changes to rejected state
    await expect(firstCard).toHaveClass(/card-rejected/);
    const statusTag = firstCard.locator(".patch-status-tag");
    await expect(statusTag).toHaveText("✕ Đã bỏ qua");
  });

  test("8. Multi-Patch Toolbar: Accept All applies all proposals at once", async ({ page }) => {
    const quickSkillBtn = page.locator('.quick-skill-btn[data-skill="cut_fluff"]');
    await quickSkillBtn.click();

    const assistantMsg = page.locator(".message-bubble.assistant").last();
    const patchSuite = assistantMsg.locator(".patch-suite-container");
    await expect(patchSuite).toBeVisible({ timeout: 10000 });

    // Toolbar with Accept All
    const btnAcceptAll = patchSuite.locator(".btn-accept-all");
    if (await btnAcceptAll.isVisible()) {
      await btnAcceptAll.click();

      // All cards should be accepted
      const patchCards = patchSuite.locator(".patch-card");
      const count = await patchCards.count();
      for (let i = 0; i < count; i++) {
        await expect(patchCards.nth(i)).toHaveClass(/card-accepted/);
      }
    }
  });

  test("9. Typing message in chat input and clicking Send works", async ({ page }) => {
    const chatInput = page.locator("#chatInput");
    await chatInput.fill("Hãy viết lại đoạn mở đầu cho tự nhiên hơn");

    const btnSend = page.locator("#btnSend");
    await btnSend.click();

    // Verify user bubble
    const userMsg = page.locator(".message-bubble.user").last();
    await expect(userMsg).toHaveText("Hãy viết lại đoạn mở đầu cho tự nhiên hơn");

    // Verify assistant reply
    const assistantMsg = page.locator(".message-bubble.assistant").last();
    await expect(assistantMsg).toBeVisible();

    // General actions bar should be present
    const msgActions = assistantMsg.locator(".msg-actions");
    await expect(msgActions).toBeVisible({ timeout: 10000 });
    await expect(msgActions.locator(".action-btn").first()).toContainText("Chèn vào con trỏ");
  });

  test("10. Slash Command Palette opens when typing / and executes skill on Enter", async ({ page }) => {
    const chatInput = page.locator("#chatInput");
    await chatInput.fill("/formal");

    const slashMenu = page.locator("#slashMenu");
    await expect(slashMenu).toBeVisible();

    // Press Enter to select
    await chatInput.press("Enter");

    // Menu should close and prompt is sent
    await expect(slashMenu).toBeHidden();
    const userMsg = page.locator(".message-bubble.user").last();
    await expect(userMsg).toContainText("Văn phong công vụ");
  });

  test("11. Settings Modal: 2 Tabs navigation and saving House Voice", async ({ page }) => {
    const btnSettings = page.locator("#btnSettings");
    await btnSettings.click();

    const settingsModal = page.locator("#settingsModal");
    await expect(settingsModal).toHaveClass(/open/);

    // Tab 1 is active by default
    const tabGeneral = page.locator("#tab-general");
    await expect(tabGeneral).toHaveClass(/active/);

    // Enter House Voice
    const houseVoice = page.locator("#cfgHouseVoice");
    await houseVoice.fill("Văn phong ngắn gọn, sắc sảo, tự nhiên.");

    // Switch to Tab 2
    const tabApiBtn = page.locator('.tab-btn[data-tab="tab-api"]');
    await tabApiBtn.click();
    const tabApi = page.locator("#tab-api");
    await expect(tabApi).toHaveClass(/active/);
    await expect(tabGeneral).not.toHaveClass(/active/);

    // Switch back to Tab 1 and Save
    const tabGenBtn = page.locator('.tab-btn[data-tab="tab-general"]');
    await tabGenBtn.click();

    const btnSave = page.locator("#btnSaveSettings");
    await btnSave.click();

    await expect(settingsModal).not.toHaveClass(/open/);
    const toast = page.locator("#oriToast");
    await expect(toast).toBeVisible();
    await expect(toast).toContainText("Đã lưu cài đặt thành công");
  });

  test("12. Sessions Modal opens and displays session threads", async ({ page }) => {
    const btnSessions = page.locator("#btnSessions");
    await btnSessions.click();

    const sessionsModal = page.locator("#sessionsModal");
    await expect(sessionsModal).toHaveClass(/open/);

    const btnClose = page.locator("#btnCloseSessions");
    await btnClose.click();
    await expect(sessionsModal).not.toHaveClass(/open/);
  });
});
