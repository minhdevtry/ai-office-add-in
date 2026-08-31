/**
 * Word-Level Diffing Module (LCS / Myers diff for text)
 * Tinh túy so sánh văn bản từng từ một để hiển thị Diff trực quan trong UI.
 *
 * KHÔNG build OOXML ở đây nữa — insertWithTrackChanges dùng pattern
 * claude-word-addin (toggle changeTrackingMode + insert text) để Word
 * tự sinh redline marks. File này chỉ phục vụ UI preview.
 */

export function tokenizeWords(text) {
  if (!text) return [];
  // Tách theo từ và giữ lại khoảng trắng / dấu câu
  return text.match(/[\p{L}\p{N}_]+|[^\s\p{L}\p{N}_]+|\s+/gu) || [];
}

/**
 * Tính toán mảng khác biệt (Diff) ở cấp độ từ (word-level).
 * @param {string} oldStr
 * @param {string} newStr
 * @returns {Array<{ type: 'equal' | 'del' | 'ins', value: string }>}
 */
export function computeWordDiff(oldStr, newStr) {
  const oldTokens = tokenizeWords(oldStr || "");
  const newTokens = tokenizeWords(newStr || "");

  const n = oldTokens.length;
  const m = newTokens.length;

  if (n === 0 && m === 0) return [];
  if (n === 0) return [{ type: "ins", value: newStr }];
  if (m === 0) return [{ type: "del", value: oldStr }];

  // Bảng LCS tối ưu bộ nhớ
  const dp = Array.from({ length: n + 1 }, () => new Uint16Array(m + 1));

  for (let i = 1; i <= n; i++) {
    for (let j = 1; j <= m; j++) {
      if (oldTokens[i - 1] === newTokens[j - 1]) {
        dp[i][j] = dp[i - 1][j - 1] + 1;
      } else {
        dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
      }
    }
  }

  // Truy vết ngược từ cuối bảng về đầu
  let i = n;
  let j = m;
  const rawDiff = [];

  while (i > 0 || j > 0) {
    if (i > 0 && j > 0 && oldTokens[i - 1] === newTokens[j - 1]) {
      rawDiff.unshift({ type: "equal", value: oldTokens[i - 1] });
      i--;
      j--;
    } else if (j > 0 && (i === 0 || dp[i][j - 1] >= dp[i - 1][j])) {
      rawDiff.unshift({ type: "ins", value: newTokens[j - 1] });
      j--;
    } else if (i > 0 && (j === 0 || dp[i][j - 1] < dp[i - 1][j])) {
      rawDiff.unshift({ type: "del", value: oldTokens[i - 1] });
      i--;
    }
  }

  // Gom nhóm các token liên tiếp cùng loại để HTML gọn gàng
  const mergedDiff = [];
  for (const chunk of rawDiff) {
    if (
      mergedDiff.length > 0 &&
      mergedDiff[mergedDiff.length - 1].type === chunk.type
    ) {
      mergedDiff[mergedDiff.length - 1].value += chunk.value;
    } else {
      mergedDiff.push({ type: chunk.type, value: chunk.value });
    }
  }

  return mergedDiff;
}

/**
 * Render Diff thành chuỗi HTML trực quan để hiển thị trong chat bubble.
 */
export function renderDiffHtml(diffChunks) {
  if (!Array.isArray(diffChunks) || diffChunks.length === 0) return "";

  return diffChunks
    .map((chunk) => {
      const escaped = String(chunk.value)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/\n/g, "<br/>");

      if (chunk.type === "del") {
        return `<del class="diff-del">${escaped}</del>`;
      }
      if (chunk.type === "ins") {
        return `<ins class="diff-ins">${escaped}</ins>`;
      }
      return `<span class="diff-equal">${escaped}</span>`;
    })
    .join("");
}
