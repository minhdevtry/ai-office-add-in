/**
 * Text Analytics & Readability Module
 * Tinh túy phân tích chỉ số văn bản, số câu, độ dài câu trung bình và độ dễ đọc
 */

export function analyzeText(text) {
  if (!text || typeof text !== "string") {
    return {
      words: 0,
      chars: 0,
      sentences: 0,
      paragraphs: 0,
      avgWordsPerSentence: 0,
      readTimeMin: 0,
      readabilityScore: 100,
      readabilityLevel: "Dễ đọc",
      complexity: "Thấp",
    };
  }

  const cleanText = text.trim();
  const chars = cleanText.length;
  
  // Tách từ (hỗ trợ tiếng Việt và Unicode)
  const wordsArray = cleanText.split(/\s+/).filter(Boolean);
  const words = wordsArray.length;

  // Tách câu theo dấu chấm, chấm hỏi, chấm than
  const sentencesArray = cleanText
    .replace(/([.!?…])\s+/g, "$1\n")
    .split("\n")
    .map((s) => s.trim())
    .filter((s) => s.length > 2);
  const sentences = Math.max(1, sentencesArray.length);

  // Tách đoạn văn
  const paragraphsArray = cleanText.split(/\n+/).filter((p) => p.trim().length > 0);
  const paragraphs = Math.max(1, paragraphsArray.length);

  const avgWordsPerSentence = Math.round((words / sentences) * 10) / 10;
  const readTimeMin = Math.ceil(words / 200); // Tốc độ đọc trung bình 200 từ/phút

  // Tính chỉ số dễ đọc tương thích (Dựa trên độ dài câu & độ dài từ trung bình)
  // Câu dài > 25 từ làm giảm độ dễ đọc
  let score = 100 - (avgWordsPerSentence - 12) * 3;
  score = Math.max(10, Math.min(100, Math.round(score)));

  let readabilityLevel = "Dễ đọc";
  let complexity = "Thấp";

  if (score >= 80) {
    readabilityLevel = "Rất trôi chảy & Dễ đọc";
    complexity = "Thấp (Phổ thông)";
  } else if (score >= 60) {
    readabilityLevel = "Chuẩn mực & Rõ ràng";
    complexity = "Trung bình (Báo chí / Công sở)";
  } else if (score >= 40) {
    readabilityLevel = "Hơi phức tạp (Nhiều câu dài)";
    complexity = "Cao (Học thuật / Hành chính)";
  } else {
    readabilityLevel = "Rất phức tạp & Khó theo dõi";
    complexity = "Rất cao (Cần tách ngắn câu)";
  }

  return {
    words,
    chars,
    sentences,
    paragraphs,
    avgWordsPerSentence,
    readTimeMin,
    readabilityScore: score,
    readabilityLevel,
    complexity,
  };
}
