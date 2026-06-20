const CHARS_KD = [
  "ñ", "Q+Z", "sas", "aa", ")Z", "ZZ", "‘", "’", "“", "”",
  "å", "ƒ", "„", "…", "†", "‡", "ˆ", "‰", "Š", "‹",
  "¶+", "d+", "[+k", "[+", "x+", "T+", "t+", "M+", "<+", "Q+", ";+", "j+", "u+",
  "Ùk", "Ù", "ä", "–", "—", "é", "™", "=kk", "f=k",
  "à", "á", "â", "ã", "ºz", "º", "í", "{k", "{", "=", "«",
  "Nî", "Vî", "Bî", "Mî", "<î", "|", "K", "}",
  "J", "Vª", "Mª", "<ªª", "Nª", "Ø", "Ý", "nzZ", "æ", "ç", "Á", "xz", "#", ":",
  "v‚", "vks", "vkS", "vk", "v", "b±", "Ã", "bZ", "b", "m", "Å", ",s", ",", "_",
  "ô", "d", "Dk", "D", "[k", "[", "x", "Xk", "X", "Ä", "?k", "?", "³",
  "pkS", "p", "Pk", "P", "N", "t", "Tk", "T", ">", "÷", "¥",
  "ê", "ë", "V", "B", "ì", "ï", "M+", "<+", "M", "<", ".k", ".",
  "r", "Rk", "R", "Fk", "F", ")", "n", "/k", "èk", "/", "Ë", "è", "u", "Uk", "U",
  "i", "Ik", "I", "Q", "¶", "c", "Ck", "C", "Hk", "H", "e", "Ek", "E",
  ";", "¸", "j", "y", "Yk", "Y", "G", "o", "Ok", "O",
  "'k", "'", "\"k", "\"", "l", "Lk", "L", "g",
  "È", "z",
  "Ì", "Í", "Î", "Ï", "Ñ", "Ò", "Ó", "Ô", "Ö", "Ø", "Ù", "Ük", "Ü",
  "‚", "ks", "kS", "k", "h", "q", "w", "`", "s", "S",
  "a", "¡", "%", "W", "•", "·", "∙", "·", "~j", "~", "\\", "+", " ः",
  "^", "*", "Þ", "ß", "(", "¼", "½", "¿", "À", "¾", "A", "-", "&", "&", "Œ", "]", "~ ", "@"
];

const CHARS_UNICODE = [
  "॰", "QZ+", "sa", "a", "र्द्ध", "Z", "\"", "\"", "'", "'",
  "०", "१", "२", "३", "४", "५", "६", "७", "८", "९",
  "फ़्", "क़", "ख़", "ख़्", "ग़", "ज़्", "ज़", "ड़", "ढ़", "फ़", "य़", "ऱ", "ऩ",
  "त्त", "त्त्", "क्त", "दृ", "कृ", "न्न", "न्न्", "=k", "f=",
  "ह्न", "ह्य", "हृ", "ह्म", "ह्र", "ह्", "द्द", "क्ष", "क्ष्", "त्र", "त्र्",
  "छ्य", "ट्य", "ठ्य", "ड्य", "ढ्य", "द्य", "ज्ञ", "द्व",
  "श्र", "ट्र", "ड्र", "ढ्र", "छ्र", "क्र", "फ्र", "र्द्र", "द्र", "प्र", "प्र", "ग्र", "रु", "रू",
  "ऑ", "ओ", "औ", "आ", "अ", "ईं", "ई", "ई", "इ", "उ", "ऊ", "ऐ", "ए", "ऋ",
  "क्क", "क", "क", "क्", "ख", "ख्", "ग", "ग", "ग्", "घ", "घ", "घ्", "ङ",
  "चै", "च", "च", "च्", "छ", "ज", "ज", "ज्", "झ", "झ्", "ञ",
  "ट्ट", "ट्ठ", "ट", "ठ", "ड्ड", "ड्ढ", "ड़", "ढ़", "ड", "ढ", "ण", "ण्",
  "त", "त", "त्", "थ", "थ्", "द्ध", "द", "ध", "ध", "ध्", "ध्", "ध्", "न", "न", "न्",
  "प", "प", "प्", "फ", "फ्", "ब", "ब", "ब्", "भ", "भ्", "म", "म", "म्",
  "य", "य्", "र", "ल", "ल", "ल्", "ळ", "व", "व", "व्",
  "श", "श्", "ष", "ष्", "स", "स", "स्", "ह",
  "ीं", "्र",
  "द्द", "ट्ट", "ट्ठ", "ड्ड", "कृ", "भ", "्य", "ड्ढ", "झ्", "क्र", "त्त्", "श", "श्",
  "ॉ", "ो", "ौ", "ा", "ी", "ु", "ू", "ृ", "े", "ै",
  "ं", "ँ", "ः", "ॅ", "ऽ", "ऽ", "ऽ", "ऽ", "्र", "्", "?", "़", ":",
  "‘", "’", "“", "”", ";", "(", ")", "{", "}", "=", "।", ".", "-", "µ", "॰", ",", "् ", "/"
];

export function krutidevToUnicode(text: string): string {
  if (!text) return "";

  let processPart = text;

  // Perform character mapping replacements
  for (let i = 0; i < CHARS_KD.length; i++) {
    const findStr = CHARS_KD[i];
    const replaceStr = CHARS_UNICODE[i];
    // Replace all occurrences of findStr with replaceStr using split/join
    processPart = processPart.split(findStr).join(replaceStr);
  }

  // Code for Replacing Special glyphs

  // Glyph1: ± (reph + anusvar)
  processPart = processPart.split("±").join("Zं");

  // Glyph2: Æ
  processPart = processPart.split("Æ").join("र्f");

  // Repositioning 'f' (ि matra)
  let position_of_i = processPart.indexOf("f");
  while (position_of_i > -1) {
    const char_next_to_i = processPart.charAt(position_of_i + 1);
    const to_replace = "f" + char_next_to_i;
    const replacement = char_next_to_i + "ि";
    processPart = processPart.replace(to_replace, replacement);
    position_of_i = processPart.indexOf("f", position_of_i + 1);
  }

  // Glyph3 & Glyph4: Ç & É
  processPart = processPart.split("Ç").join("fa");
  processPart = processPart.split("É").join("र्fa");

  // Repositioning 'fa' (िं matra)
  let position_of_fa = processPart.indexOf("fa");
  while (position_of_fa > -1) {
    const char_next_to_fa = processPart.charAt(position_of_fa + 2);
    const to_replace = "fa" + char_next_to_fa;
    const replacement = char_next_to_fa + "िं";
    processPart = processPart.replace(to_replace, replacement);
    position_of_fa = processPart.indexOf("fa", position_of_fa + 1);
  }

  // Glyph5: Ê
  processPart = processPart.split("Ê").join("ीZ");

  // Eliminate wrong chhoti ee matra on half-letters (e.g. ि् -> ्ि)
  while (processPart.indexOf("ि्") > -1) {
    const pos = processPart.indexOf("ि्");
    const next_char = processPart.charAt(pos + 2);
    const to_replace = "ि्" + next_char;
    const replacement = "्" + next_char + "ि";
    processPart = processPart.replace(to_replace, replacement);
  }

  // Eliminating reph "Z" and putting 'half-r' (र्) at the proper position.
  const set_of_matras = "अ आ इ ई उ ऊ ए ऐ ओ औ ा ि ी ु ू ृ े ै ो ौ ं : ँ ॅ";
  let position_of_R = processPart.indexOf("Z");
  while (position_of_R > -1) {
    let probable_position_of_half_r = position_of_R - 1;
    let char_at_probable = processPart.charAt(probable_position_of_half_r);

    // Skip backward over matras and half-letter structures (halants)
    while (true) {
      if (set_of_matras.indexOf(char_at_probable) > -1) {
        probable_position_of_half_r--;
        char_at_probable = processPart.charAt(probable_position_of_half_r);
      } else if (probable_position_of_half_r > 0 && processPart.charAt(probable_position_of_half_r - 1) === "्") {
        // Skip halant and the preceding consonant (e.g. स् -> move back by 2)
        probable_position_of_half_r -= 2;
        char_at_probable = processPart.charAt(probable_position_of_half_r);
      } else {
        break;
      }
    }

    const char_to_be_replaced = processPart.substring(probable_position_of_half_r, position_of_R);
    const replacement_str = "र्" + char_to_be_replaced;
    const full_target = char_to_be_replaced + "Z";
    processPart = processPart.replace(full_target, replacement_str);
    position_of_R = processPart.indexOf("Z");
  }

  return processPart;
}
