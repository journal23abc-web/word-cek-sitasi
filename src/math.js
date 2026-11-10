/*
 * TplMath — rumus/persamaan untuk add-in Template Paste (tanpa Office.js).
 *   LaTeX / teks biasa / MathML  ->  AST  ->  OMML (persamaan Word asli), MathML (pratinjau), teks Unicode
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.TplMath = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  /* ---------------- simbol ---------------- */
  const GREEK = {
    alpha: "α", beta: "β", gamma: "γ", delta: "δ", epsilon: "ϵ", varepsilon: "ε", zeta: "ζ", eta: "η", theta: "θ", vartheta: "ϑ",
    iota: "ι", kappa: "κ", lambda: "λ", mu: "μ", nu: "ν", xi: "ξ", pi: "π", varpi: "ϖ", rho: "ρ", varrho: "ϱ", sigma: "σ", varsigma: "ς",
    tau: "τ", upsilon: "υ", phi: "ϕ", varphi: "φ", chi: "χ", psi: "ψ", omega: "ω",
    Gamma: "Γ", Delta: "Δ", Theta: "Θ", Lambda: "Λ", Xi: "Ξ", Pi: "Π", Sigma: "Σ", Upsilon: "Υ", Phi: "Φ", Psi: "Ψ", Omega: "Ω",
  };
  const OPS = {
    times: "×", cdot: "⋅", pm: "±", mp: "∓", div: "÷", ast: "∗", star: "⋆", circ: "∘", bullet: "∙",
    leq: "≤", le: "≤", geq: "≥", ge: "≥", neq: "≠", ne: "≠", approx: "≈", sim: "∼", simeq: "≃", cong: "≅", equiv: "≡", propto: "∝",
    ll: "≪", gg: "≫", to: "→", rightarrow: "→", leftarrow: "←", Rightarrow: "⇒", Leftarrow: "⇐", leftrightarrow: "↔", Leftrightarrow: "⇔", mapsto: "↦",
    in: "∈", notin: "∉", ni: "∋", subset: "⊂", subseteq: "⊆", supset: "⊃", supseteq: "⊇", cup: "∪", cap: "∩", setminus: "∖",
    forall: "∀", exists: "∃", neg: "¬", land: "∧", lor: "∨", wedge: "∧", vee: "∨", mid: "∣", parallel: "∥", perp: "⊥",
    infty: "∞", partial: "∂", nabla: "∇", emptyset: "∅", varnothing: "∅", degree: "°", prime: "′", angle: "∠", therefore: "∴", because: "∵",
    ldots: "…", cdots: "⋯", vdots: "⋮", ddots: "⋱", dots: "…", hbar: "ℏ", ell: "ℓ", Re: "ℜ", Im: "ℑ", aleph: "ℵ",
    langle: "⟨", rangle: "⟩", lvert: "|", rvert: "|", lVert: "‖", rVert: "‖", vert: "|", Vert: "‖", lfloor: "⌊", rfloor: "⌋", lceil: "⌈", rceil: "⌉",
    "{": "{", "}": "}", "%": "%", "$": "$", "&": "&", "#": "#", "_": "_", "|": "‖",
    odot: "⊙", oplus: "⊕", otimes: "⊗", ominus: "⊖", oslash: "⊘", top: "⊤", bot: "⊥", intercal: "⊺", dagger: "†", ddagger: "‡",
    triangleq: "≜", coloneqq: "≔", doteq: "≐", lesssim: "≲", gtrsim: "≳", leqslant: "⩽", geqslant: "⩾", nleq: "≰", ngeq: "≱",
    sqcup: "⊔", sqcap: "⊓", uplus: "⊎", wr: "≀", diamond: "⋄", triangle: "△", square: "□", checkmark: "✓", colon: ":",
    longrightarrow: "⟶", longleftarrow: "⟵", Longrightarrow: "⟹", implies: "⟹", iff: "⟺", uparrow: "↑", downarrow: "↓", hookrightarrow: "↪",
    leftarrow: "←", gets: "←", lbrace: "{", rbrace: "}", lbrack: "[", rbrack: "]", backslash: "∖", ast: "∗", S: "§", P: "¶", copyright: "©",
    varkappa: "ϰ", digamma: "ϝ", imath: "ı", jmath: "ȷ", wp: "℘", mho: "℧", complement: "∁", nexists: "∄", sharp: "♯", flat: "♭", natural: "♮",
  };
  const NARY = { sum: "∑", prod: "∏", coprod: "∐", int: "∫", iint: "∬", iiint: "∭", oint: "∮", bigcup: "⋃", bigcap: "⋂" };
  const FUNCS = new Set(["log", "ln", "exp", "sin", "cos", "tan", "cot", "sec", "csc", "arcsin", "arccos", "arctan", "sinh", "cosh", "tanh",
    "max", "min", "lim", "sup", "inf", "det", "dim", "ker", "deg", "gcd", "arg", "Pr", "var", "Var", "cov", "Cov", "corr", "E", "argmax", "argmin", "sgn", "mod"]);
  const TEX_FUNCS = new Set(["log", "ln", "exp", "sin", "cos", "tan", "cot", "sec", "csc", "arcsin", "arccos", "arctan", "sinh", "cosh", "tanh", "max", "min", "lim", "sup", "inf", "det", "dim", "ker", "deg", "gcd", "arg", "Pr"]);
  const LIMFUNCS = new Set(["lim", "max", "min", "sup", "inf", "argmax", "argmin"]);
  const ACCENTS = { hat: "\u0302", widehat: "\u0302", bar: "\u0305", overline: "\u0305", tilde: "\u0303", widetilde: "\u0303", dot: "\u0307", ddot: "\u0308", vec: "\u20D7", check: "\u030C", acute: "\u0301", grave: "\u0300", breve: "\u0306" };
  const SPACES = { ",": " ", ";": " ", ":": " ", quad: "  ", qquad: "    ", " ": " ", "!": "", enspace: " ", thinspace: " " };
  const TEXTCMD = new Set(["text", "textrm", "mathrm", "operatorname", "textit", "mbox", "textup", "texttt", "textsf", "begin", "end"]);
  const SCRIPTS = { mathbb: "double-struck", mathcal: "script", mathscr: "script", mathfrak: "fraktur", mathsf: "sans-serif", mathtt: "monospace", Bbb: "double-struck" };
  // huruf matematika Unicode (untuk teks & pratinjau): ℝ, 𝒟, 𝔉, …
  const ALNUM_BASE = { bold: [0x1D400, 0x1D41A], "double-struck": [0x1D538, 0x1D552], script: [0x1D49C, 0x1D4B6], fraktur: [0x1D504, 0x1D51E], "sans-serif": [0x1D5A0, 0x1D5BA], monospace: [0x1D670, 0x1D68A] };
  const ALNUM_EXC = {
    "double-struck": { C: "ℂ", H: "ℍ", N: "ℕ", P: "ℙ", Q: "ℚ", R: "ℝ", Z: "ℤ" },
    script: { B: "ℬ", E: "ℰ", F: "ℱ", H: "ℋ", I: "ℐ", L: "ℒ", M: "ℳ", R: "ℛ", e: "ℯ", g: "ℊ", o: "ℴ" },
    fraktur: { C: "ℭ", H: "ℌ", I: "ℑ", R: "ℜ", Z: "ℨ" },
  };
  const styledChar = (ch, scr) => {
    const ex = ALNUM_EXC[scr] && ALNUM_EXC[scr][ch];
    if (ex) return ex;
    const b = ALNUM_BASE[scr];
    if (!b) return ch;
    if (/[A-Z]/.test(ch)) return String.fromCodePoint(b[0] + ch.charCodeAt(0) - 65);
    if (/[a-z]/.test(ch)) return String.fromCodePoint(b[1] + ch.charCodeAt(0) - 97);
    return ch;
  };
  const SUBDIG = { "₀": "0", "₁": "1", "₂": "2", "₃": "3", "₄": "4", "₅": "5", "₆": "6", "₇": "7", "₈": "8", "₉": "9", "₊": "+", "₋": "-", "ₐ": "a", "ₑ": "e", "ᵢ": "i", "ⱼ": "j", "ₖ": "k", "ₙ": "n", "ₜ": "t", "ₓ": "x" };
  const SUPDIG = { "⁰": "0", "¹": "1", "²": "2", "³": "3", "⁴": "4", "⁵": "5", "⁶": "6", "⁷": "7", "⁸": "8", "⁹": "9", "⁺": "+", "⁻": "-", "ⁿ": "n", "ⁱ": "i" };
  const TO_SUB = Object.fromEntries(Object.entries(SUBDIG).map(([k, v]) => [v, k]));
  const TO_SUP = Object.fromEntries(Object.entries(SUPDIG).map(([k, v]) => [v, k]));

  /* ---------------- input: deteksi & normalisasi ---------------- */
  const looksLatex = (s) => /\\[a-zA-Z]+|\\[,;!{}]|[\^_]\{|\$/.test(s);

  /**
   * Bersihkan masukan: buang $$, \[ \], \begin{equation}, \label, ambil \tag{n} / "(n)" di akhir sebagai nomor.
   * Teks biasa (β₀, X1, x²) diubah ke LaTeX.
   * return { tex, number }
   */
  function normalizeInput(raw, opts) {
    opts = opts || {};
    let s = repairLatex(String(raw || "").replace(/\r/g, "").trim());
    let number = null;
    s = s.replace(/^\$\$([\s\S]*)\$\$$/, "$1").replace(/^\\\[([\s\S]*)\\\]$/, "$1").replace(/^\\\(([\s\S]*)\\\)$/, "$1").replace(/^\$([\s\S]*)\$$/, "$1");
    s = s.replace(/\\begin\{(equation|align|gather|displaymath)\*?\}([\s\S]*)\\end\{\1\*?\}/, "$2");
    s = s.replace(/\\label\{[^}]*\}/g, "").replace(/\\nonumber|\\notag/g, "");
    const tag = /\\tag\{([^}]*)\}/.exec(s);
    if (tag) { number = tag[1]; s = s.replace(tag[0], ""); }
    const tail = /\s+\((\d+[a-z]?)\)\s*$/.exec(s);
    if (tail && !/[=+\-]\s*\((\d+[a-z]?)\)\s*$/.test(s)) { number = tail[1]; s = s.slice(0, tail.index); }
    if (/\\\\/.test(s) && !/\\begin\{/.test(s.replace(/\\\\\s*$/, ""))) s = "\\begin{aligned}" + s.replace(/\\\\\s*$/, "") + "\\end{aligned}";
    s = s.trim();
    if (opts.forceLinear || !looksLatex(s)) s = linearToLatex(s, opts);
    return { tex: s.trim(), number };
  }

  /**
   * Perbaiki LaTeX yang backslash-nya hilang saat disalin (mis. dari ChatGPT ke Word):
   *   "\sigma!\left("  (aslinya \!\left)  -> "\sigma\left("
   *   "…,\" di akhir baris (aslinya \\ ganti baris) -> "…,\\"
   *   "[" / "]" pembungkus rumus baris (aslinya \[ \]) -> dibuang
   */
  /** Huruf matematika Unicode (𝑖 𝑡 𝜀 𝐗 ℎ — hasil salin rumus Word sebagai teks) -> huruf biasa. */
  const mathAlnumToPlain = (s) => String(s).replace(/[\u{1D400}-\u{1D7FF}\u210E]/gu, (c) => c.normalize("NFKC"));
  function repairLatex(s) {
    let t = mathAlnumToPlain(String(s));
    if (/^\[\s*\n[\s\S]*\n\s*\]$/.test(t)) t = t.replace(/^\[\s*\n/, "").replace(/\n\s*\]$/, "");
    else if (/^\[\s[\s\S]*\\[a-zA-Z][\s\S]*\s\]$/.test(t)) t = t.replace(/^\[\s+/, "").replace(/\s+\]$/, "");
    t = t.replace(/(^|[^\\])!(?=\\(?:left|right|big|Big|bigg|Bigg)\b)/g, "$1");
    t = t.replace(/(^|[^\\])\\[ \t]*\n/g, "$1\\\\\n");
    t = t.replace(/(\d{3,}),(?=\\?[A-Za-z])/g, "$1\\,"); // "1000,t" (aslinya 1000\,t)
    // di dalam aligned/cases/matrix tanpa "\\" sama sekali: "\ " adalah ganti baris yang kehilangan satu backslash
    t = t.replace(/(\\begin\{(aligned|align\*?|gathered|cases|[pbvBV]?matrix|array)\}(?:\{[^}]*\})?)([\s\S]*?)(\\end\{\2\})/g, (m, a, env, body, z) =>
      /\\\\/.test(body) ? m : a + body.replace(/(^|[^\\])\\(?=\s)/g, "$1\\\\") + z);
    return t;
  }

  /** Teks biasa -> LaTeX: subskrip/superskrip Unicode, X1 -> X_{1}, x^2, sqrt(x), sum -> \sum */
  function linearToLatex(s, opts) {
    opts = opts || {};
    let t = s;
    t = t.replace(/([₀-₉₊₋ₐₑᵢⱼₖₙₜₓ]+)/g, (m) => "_{" + [...m].map((c) => SUBDIG[c]).join("") + "}");
    t = t.replace(/([⁰-⁹¹²³⁺⁻ⁿⁱ]+)/g, (m) => "^{" + [...m].map((c) => SUPDIG[c]).join("") + "}");
    t = t.replace(/×\s*10\s*([−-])\s*(\d+)(?![\d.])/g, "×10^{-$2}"); // 1.58×10−7 (superskrip hilang saat disalin)
    // indeks data panel tanpa format subskrip: Yit, SHIPit, εit -> Y_{it}, SHIP_{it}
    t = t.replace(/(^|[^A-Za-z\\])([A-Z]{1,15}|[α-ωΑ-Ω])(it|jt|ijt|ij|ict)(?![A-Za-z])/g, "$1$2_{$3}");
    t = t.replace(/(^|[^\\A-Za-z{])([A-Z]{2})(?=_\{)/g, "$1\\mathrm{$2}"); // EG_{it}, DI_{it} tegak seperti SHIP_{it}
    t = t.replace(/√\s*\(([^()]*)\)/g, "\\sqrt{$1}").replace(/\bsqrt\s*\(([^()]*)\)/g, "\\sqrt{$1}").replace(/√\s*(\w+)/g, "\\sqrt{$1}");
    t = t.replace(/Σ|∑/g, "\\sum ").replace(/∏/g, "\\prod ");
    // kata/singkatan >= 3 huruf (ROA, DER, Var, log) -> tegak; fungsi dikenal -> \\log dst.
    t = t.replace(/(^|[^\\A-Za-z])([A-Z]{3,}|[a-z]{3,}|[A-Z][a-z]{2,}|ln)(?![A-Za-z])/g, (m, pre, w) => pre + (FUNCS.has(w) ? "\\" + w + " " : /^[a-z]+$/.test(w) && w.length > 3 ? "\\text{" + (pre === " " ? " " : "") + w + "}" : "\\mathrm{" + w + "}"));
    if (opts.autoSubscript !== false) { // X1 -> X_{1}; dua kali karena pola tidak tumpang-tindih (β1X1)
      // X1 -> X_{1}; X1c (variabel terpusat) -> X_{1c}
      const re = /(^|[^A-Za-z\\{_^])([A-Za-zα-ωΑ-Ω])(\d+(?:it|jt|ijt|[a-z](?![a-z{]))?)(?![\d.])/g;
      t = t.replace(re, "$1$2_{$3}").replace(re, "$1$2_{$3}");
    }
    t = t.replace(/\*/g, "\\cdot ").replace(/<=|≤/g, "\\leq ").replace(/>=|≥/g, "\\geq ").replace(/!=|≠/g, "\\neq ");
    t = t.replace(/\^(\w+|\([^()]*\))/g, (m, x) => "^{" + x.replace(/^\((.*)\)$/, "$1") + "}");
    // a/b, (a+b)/(c+d) -> pecahan bertingkat (seperti format linear Word)
    const opd = "(\\([^()]*\\)|(?:\\\\mathrm\\{[^}]*\\}|[A-Za-z0-9.α-ωΑ-Ω]|_\\{[^}]*\\}|\\^\\{[^}]*\\})+)";
    const reFrac = new RegExp(opd + "\\s*/\\s*" + opd);
    const unp = (x) => (/^\(.*\)$/.test(x) ? x.slice(1, -1) : x);
    for (let k = 0; k < 8 && reFrac.test(t); k++) t = t.replace(reFrac, (m, a, b) => "\\frac{" + unp(a) + "}{" + unp(b) + "}");
    return t;
  }

  /* ---------------- tokenizer & parser LaTeX ---------------- */
  function tokenize(src) {
    const toks = [];
    let i = 0;
    while (i < src.length) {
      const c = src[i];
      if (c === "\\") {
        const m = /^\\([a-zA-Z]+)\*?/.exec(src.slice(i));
        if (m) {
          toks.push({ k: "cmd", v: m[1] }); i += m[0].length;
          if (TEXTCMD.has(m[1])) { // isi \text{...} diambil mentah agar spasi tetap
            let j = i; while (src[j] === " ") j++;
            if (src[j] === "{") {
              let d = 1, k = j + 1;
              while (k < src.length && d) { if (src[k] === "{") d++; else if (src[k] === "}") d--; if (d) k++; }
              toks.push({ k: "raw", v: src.slice(j + 1, k).replace(/\\([%$&#_{} ])/g, "$1") }); i = k + 1;
            }
          }
          continue;
        }
        toks.push({ k: "cmd", v: src[i + 1] || "" }); i += 2; continue;
      }
      if (c === "{" || c === "}") { toks.push({ k: c }); i++; continue; }
      if (c === "^" || c === "_" || c === "&") { toks.push({ k: c }); i++; continue; }
      if (/\s/.test(c)) { i++; continue; }
      if (/[0-9]/.test(c)) { const m = /^[0-9]+(?:\.[0-9]+)?/.exec(src.slice(i)); toks.push({ k: "num", v: m[0] }); i += m[0].length; continue; }
      // karakter di luar BMP (mis. 𝑖 𝜀 dari rumus Word) = 2 unit UTF-16: jangan dipotong
      const ch = String.fromCodePoint(src.codePointAt(i));
      if (/[\p{L}]/u.test(ch)) { toks.push({ k: "id", v: ch }); i += ch.length; continue; }
      if (ch === "'") { toks.push({ k: "op", v: "′" }); i++; continue; }
      if (/[\p{M}]/u.test(ch) && toks.length) { toks[toks.length - 1].v += ch; i += ch.length; continue; } // tanda diakritik gabung
      toks.push({ k: "op", v: ch === "-" ? "−" : ch }); i += ch.length;
    }
    return toks;
  }

  function parseLatex(src) {
    const toks = tokenize(src);
    let p = 0;
    const peek = () => toks[p];
    const next = () => toks[p++];
    const row = (items) => ({ t: "row", items });

    function readGroupRaw() { // isi {...} sebagai string mentah (untuk \text)
      if (peek() && peek().k === "raw") return next().v;
      if (!peek() || peek().k !== "{") return "";
      next(); let depth = 1, out = "";
      while (p < toks.length) {
        const tk = next();
        if (tk.k === "{") { depth++; out += "{"; continue; }
        if (tk.k === "}") { depth--; if (!depth) break; out += "}"; continue; }
        out += tk.k === "cmd" ? (OPS[tk.v] || GREEK[tk.v] || (SPACES[tk.v] != null ? " " : "")) : tk.v != null ? tk.v : "";
        if (tk.k === "num" || tk.k === "id") out += "";
      }
      return out;
    }
    function parseArg() { // satu argumen: {grup} atau satu token
      const tk = peek();
      if (!tk) return row([]);
      if (tk.k === "{") { next(); const r = parseRow("}"); if (peek() && peek().k === "}") next(); return r; }
      return parseAtom() || row([]);
    }
    function parseAtom() {
      const tk = next();
      if (!tk) return null;
      if (tk.k === "num") return { t: "num", v: tk.v };
      if (tk.k === "id") return { t: "id", v: tk.v };
      if (tk.k === "op") {
        if (tk.v === "(" || tk.v === "[") { // kurung biasa -> delimiter
          const close = tk.v === "(" ? ")" : "]";
          const save = p; const inner = parseRow(close);
          if (peek() && peek().k === "op" && peek().v === close) { next(); return { t: "delim", open: tk.v, close, e: inner }; }
          p = save; return { t: "op", v: tk.v };
        }
        return { t: "op", v: tk.v };
      }
      if (tk.k === "{") {
        const r = parseRow("}"); if (peek() && peek().k === "}") next();
        // {,} = koma tanpa spasi (5{,}898); grup lepas berisi koma = himpunan {a, b} (backslash \{ hilang saat disalin)
        if (r.items.length === 1 && r.items[0].t === "op" && r.items[0].v === ",") return r.items[0];
        if (r.items.some((x) => x.t === "op" && x.v === ",") ||
            (r.items.length === 1 && r.items[0].t === "delim" && r.items[0].open === "(" && /,/.test(toUnicode(r.items[0].e)) && peek() && (peek().k === "_" || peek().k === "^")))
          return { t: "delim", open: "{", close: "}", e: r };
        return r;
      }
      if (tk.k === "}") return null;
      if (tk.k === "cmd") {
        const c = tk.v;
        if (c === "frac" || c === "dfrac" || c === "tfrac" || c === "cfrac") { const n = parseArg(), d = parseArg(); return { t: "frac", n, d }; }
        if (c === "binom" || c === "tbinom" || c === "dbinom") { const n = parseArg(), d = parseArg(); return { t: "delim", open: "(", close: ")", e: { t: "frac", n, d, noBar: true } }; }
        if (c === "sqrt") {
          let deg = null;
          if (peek() && peek().k === "op" && peek().v === "[") { next(); deg = parseRow("]"); if (peek() && peek().v === "]") next(); }
          return { t: "sqrt", e: parseArg(), deg };
        }
        if (c === "text" || c === "textrm" || c === "mathrm" || c === "operatorname" || c === "textit" || c === "mbox" || c === "textup" || c === "texttt" || c === "textsf") {
          const raw = readGroupRaw();
          if (c === "operatorname" || c === "mathrm") return { t: "fn", v: raw };
          return { t: "text", v: raw };
        }
        if (c === "mathbf" || c === "boldsymbol" || c === "bm" || c === "textbf") { const a = parseArg(); return { t: "style", sty: "b", e: a }; }
        if (c === "mathit") return parseArg();
        if (SCRIPTS[c]) return { t: "style", scr: SCRIPTS[c], e: parseArg() };
        if (ACCENTS[c]) return { t: "acc", chr: ACCENTS[c], e: parseArg(), name: c };
        if (c === "left") {
          let open = next(); open = open ? (open.k === "cmd" ? (OPS[open.v] || (open.v === "." ? "" : open.v)) : open.v) : "(";
          const inner = parseRow("\\right");
          let close = "";
          if (peek() && peek().k === "cmd" && peek().v === "right") { next(); const c2 = next(); close = c2 ? (c2.k === "cmd" ? (OPS[c2.v] || (c2.v === "." ? "" : c2.v)) : c2.v) : ")"; }
          return { t: "delim", open, close, e: inner };
        }
        if (c === "right") return null;
        if (c === "big" || c === "Big" || c === "bigg" || c === "Bigg" || c === "bigl" || c === "bigr" || c === "Bigl" || c === "Bigr" || c === "displaystyle" || c === "textstyle" || c === "limits" || c === "nolimits") return { t: "row", items: [] };
        if (NARY[c]) return { t: "nary", chr: NARY[c], name: c };
        if (FUNCS.has(c)) return { t: "fn", v: c };
        if (GREEK[c]) return { t: "id", v: GREEK[c] };
        if (OPS[c]) return { t: "op", v: OPS[c] };
        if (SPACES[c] != null) return { t: "space", v: SPACES[c] };
        if (c === "begin") { // lingkungan: cases / matrix / aligned -> baris-baris
          const env = readGroupRaw().replace(/\*$/, "");
          const rows = parseEnv();
          if (/cases/.test(env)) return { t: "delim", open: "{", close: "", e: { t: "eqarr", rows: rows.map((r) => joinCells(r, "  ")) } };
          if (/matrix|array/.test(env)) {
            const m = { t: "matrix", rows };
            const br = { pmatrix: ["(", ")"], bmatrix: ["[", "]"], Bmatrix: ["{", "}"], vmatrix: ["|", "|"], Vmatrix: ["‖", "‖"] }[env];
            return br ? { t: "delim", open: br[0], close: br[1], e: m } : m;
          }
          // aligned/split: "&" = titik rata (m:aln) agar baris-baris sejajar di tanda "="
          return { t: "eqarr", rows: rows.map((r) => ({ t: "row", items: r.reduce((a, c, i) => a.concat(i ? [{ t: "aln" }] : [], [c]), []) })) };
        }
        if (c === "\\") return { t: "space", v: " " };
        return { t: "text", v: c }; // perintah tak dikenal: tampilkan namanya
      }
      if (tk.k === "&") return { t: "space", v: " " };
      return null;
    }
    function withScripts(base) {
      let sub = null, sup = null;
      for (;;) {
        const tk = peek();
        if (!tk) break;
        if (tk.k === "_" && !sub) { next(); sub = parseArg(); continue; }
        if (tk.k === "^" && !sup) { next(); sup = parseArg(); continue; }
        if (tk.k === "op" && tk.v === "′" && !sup) { next(); sup = { t: "op", v: "′" }; continue; }
        break;
      }
      if (base && base.t === "nary") return Object.assign(base, { sub, sup, e: null });
      if (base && base.t === "fn" && LIMFUNCS.has(base.v) && sub && !sup) return { t: "limlow", fn: base, lim: sub };
      if (sub && sup) return { t: "subsup", base, sub, sup };
      if (sub) return { t: "sub", base, sub };
      if (sup) return { t: "sup", base, sup };
      return base;
    }
    const joinCells = (cells, gap) => ({ t: "row", items: cells.reduce((a, c, i) => a.concat(i && gap ? [{ t: "space", v: gap }] : [], [c]), []) });
    function parseEnv() { // isi \begin..\end -> [[sel, ...], ...]
      const rows = [[]];
      while (p < toks.length) {
        rows[rows.length - 1].push(parseRow("env"));
        const tk = peek();
        if (!tk) break;
        if (tk.k === "&") { next(); continue; }
        if (tk.k === "cmd" && tk.v === "\\") { next(); rows.push([]); continue; }
        if (tk.k === "cmd" && tk.v === "end") { next(); readGroupRaw(); break; }
        next();
      }
      return rows.filter((r) => r.some((c) => c.items && c.items.length));
    }
    function parseRow(stop) {
      const items = [];
      while (p < toks.length) {
        const tk = peek();
        if (stop === "}" && tk.k === "}") break;
        if (stop === ")" && tk.k === "op" && tk.v === ")") break;
        if (stop === "]" && tk.k === "op" && tk.v === "]") break;
        if (stop === "\\right" && tk.k === "cmd" && tk.v === "right") break;
        if (stop === "\\end" && tk.k === "cmd" && tk.v === "end") break;
        if (stop === "env" && (tk.k === "&" || (tk.k === "cmd" && (tk.v === "\\" || tk.v === "end")))) break;
        if (!stop && tk.k === "}") { next(); continue; }
        const a = parseAtom();
        if (!a) { if (tk.k === "}" && stop !== "}") continue; break; }
        items.push(withScripts(a));
      }
      // operator n-ary: badan = elemen berikutnya (atau sisa baris bila elemen berikutnya kurung/huruf)
      // fungsi (log, sin, …) mengikat argumen berikutnya — sebelum badan ∑ dikumpulkan
      for (let k = 0; k < items.length; k++) {
        const it = items[k];
        if (it.t === "fn" && !LIMFUNCS.has(it.v) && !it.arg && k + 1 < items.length && items[k + 1].t !== "op" && items[k + 1].t !== "space") it.arg = items.splice(k + 1, 1)[0];
      }
      for (let k = 0; k < items.length; k++) {
        const it = items[k];
        if (it.t === "nary" && it.e === null) {
          const body = [];
          while (k + 1 < items.length && !(items[k + 1].t === "op" && /^[=+−±<>≤≥≈,]$/.test(items[k + 1].v))) body.push(items.splice(k + 1, 1)[0]);
          it.e = row(body);
        }
        if (it.t === "limlow" && k + 1 < items.length && items[k + 1].t !== "op") it.arg = items.splice(k + 1, 1)[0];
      }
      return row(items);
    }
    return parseRow(null);
  }

  /* ---------------- MathML (dari halaman web) -> AST ---------------- */
  function fromMathML(el) {
    const ann = el.querySelector && el.querySelector('annotation[encoding="application/x-tex"], annotation[encoding="TeX"]');
    if (ann && ann.textContent.trim()) return { tex: ann.textContent.trim() };
    const kids = (e) => Array.from(e.children || []).filter((x) => !/^(annotation|annotation-xml)$/i.test(x.localName || x.tagName));
    function conv(e) {
      const n = (e.localName || e.tagName || "").toLowerCase();
      const k = kids(e);
      const txt = (e.textContent || "").trim();
      switch (n) {
        case "math": case "mrow": case "mstyle": case "semantics": case "mpadded": case "mphantom": return { t: "row", items: k.map(conv) };
        case "mi": return txt.length > 1 ? { t: "fn", v: txt } : { t: "id", v: txt };
        case "mn": return { t: "num", v: txt };
        case "mo": return NARY_CHARS.has(txt) ? { t: "nary", chr: txt, e: { t: "row", items: [] } } : { t: "op", v: txt };
        case "mtext": case "ms": return { t: "text", v: e.textContent };
        case "mspace": return { t: "space", v: " " };
        case "mfrac": return { t: "frac", n: conv(k[0]), d: conv(k[1]) };
        case "msqrt": return { t: "sqrt", e: { t: "row", items: k.map(conv) } };
        case "mroot": return { t: "sqrt", e: conv(k[0]), deg: conv(k[1]) };
        case "msup": return { t: "sup", base: conv(k[0]), sup: conv(k[1]) };
        case "msub": return { t: "sub", base: conv(k[0]), sub: conv(k[1]) };
        case "msubsup": return { t: "subsup", base: conv(k[0]), sub: conv(k[1]), sup: conv(k[2]) };
        case "munderover": { const b = conv(k[0]); if (b.t === "nary") return Object.assign(b, { sub: conv(k[1]), sup: conv(k[2]) }); return { t: "subsup", base: b, sub: conv(k[1]), sup: conv(k[2]) }; }
        case "munder": { const b = conv(k[0]); if (b.t === "nary") return Object.assign(b, { sub: conv(k[1]) }); if (b.t === "fn") return { t: "limlow", fn: b, lim: conv(k[1]) }; return { t: "sub", base: b, sub: conv(k[1]) }; }
        case "mover": { const b = conv(k[0]), a = (k[1] && k[1].textContent || "").trim(); const ch = { "^": "\u0302", "ˆ": "\u0302", "¯": "\u0305", "‾": "\u0305", "~": "\u0303", "˜": "\u0303", "→": "\u20D7", "˙": "\u0307", ".": "\u0307" }[a]; return ch ? { t: "acc", chr: ch, e: b } : { t: "sup", base: b, sup: conv(k[1]) }; }
        case "mfenced": return { t: "delim", open: e.getAttribute("open") || "(", close: e.getAttribute("close") || ")", e: { t: "row", items: k.map(conv) } };
        case "mtable": return { t: "row", items: k.map(conv) };
        case "mtr": case "mtd": return { t: "row", items: k.map(conv).concat([{ t: "space", v: " " }]) };
        default: return k.length ? { t: "row", items: k.map(conv) } : { t: "text", v: txt };
      }
    }
    const ast = conv(el);
    // pasangkan operator n-ary dengan badannya
    (function fix(node) {
      if (!node || typeof node !== "object") return;
      if (node.t === "row") {
        for (let k = 0; k < node.items.length; k++) {
          const it = node.items[k];
          if (it.t === "nary" && (!it.e || !it.e.items || !it.e.items.length)) {
            const body = [];
            while (k + 1 < node.items.length && !(node.items[k + 1].t === "op" && /^[=+−±<>≤≥≈,]$/.test(node.items[k + 1].v))) body.push(node.items.splice(k + 1, 1)[0]);
            it.e = { t: "row", items: body };
          }
        }
      }
      Object.values(node).forEach((v) => { if (Array.isArray(v)) v.forEach(fix); else if (v && typeof v === "object") fix(v); });
    })(ast);
    return { ast };
  }
  const NARY_CHARS = new Set(["∑", "∏", "∐", "∫", "∬", "∭", "∮", "⋃", "⋂"]);

  /* ---------------- keluaran: OMML ---------------- */
  const xe = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  function ommlRun(text, o) {
    o = o || {};
    let rpr = "";
    // skema OMML: <m:nor/> ATAU (<m:scr/>, <m:sty/>) — tidak boleh keduanya (Word menolak dengan GeneralException)
    if (o.nor) rpr += "<m:nor/>";
    else {
      if (o.scr) rpr += `<m:scr m:val="${o.scr}"/>`;
      if (o.sty || o.scr) rpr += `<m:sty m:val="${o.sty || "p"}"/>`;
    }
    if (o._st && o._st.aln) { rpr += "<m:aln/>"; o._st.aln = false; }
    const w = `<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/>${o.size ? `<w:sz w:val="${Math.round(o.size * 2)}"/><w:szCs w:val="${Math.round(o.size * 2)}"/>` : ""}</w:rPr>`;
    return `<m:r>${rpr ? `<m:rPr>${rpr}</m:rPr>` : ""}${w}<m:t xml:space="preserve">${xe(text)}</m:t></m:r>`;
  }
  function toOMML(node, o) {
    o = o || {};
    if (!o._st) o = Object.assign({}, o, { _st: {} });
    const R = (n) => toOMML(n, o);
    const E = (n) => `<m:e>${R(n)}</m:e>`;
    if (!node) return "";
    switch (node.t) {
      case "row": return node.items.map(R).join("");
      case "id": return ommlRun(node.v, Object.assign({}, o, o.sty ? {} : {}));
      case "num": return ommlRun(node.v, o);
      case "op": return ommlRun(node.v, o);
      case "space": return node.v ? ommlRun(node.v, o) : "";
      case "text": return ommlRun(node.v, Object.assign({}, o, { nor: true, sty: null, scr: null }));
      case "fn": {
        const name = ommlRun(node.v, Object.assign({}, o, { sty: "p" }));
        if (node.arg) return `<m:func><m:fName>${name}</m:fName>${E(node.arg)}</m:func>`;
        return name;
      }
      case "style": return toOMML(node.e, Object.assign({}, o, node.scr ? { scr: node.scr } : { sty: node.sty }));
      case "frac": return `<m:f>${node.noBar ? '<m:fPr><m:type m:val="noBar"/></m:fPr>' : ""}<m:num>${R(node.n)}</m:num><m:den>${R(node.d)}</m:den></m:f>`;
      case "sqrt": return node.deg ? `<m:rad><m:deg>${R(node.deg)}</m:deg>${E(node.e)}</m:rad>` : `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/>${E(node.e)}</m:rad>`;
      case "sup": return `<m:sSup>${E(node.base)}<m:sup>${R(node.sup)}</m:sup></m:sSup>`;
      case "sub": return `<m:sSub>${E(node.base)}<m:sub>${R(node.sub)}</m:sub></m:sSub>`;
      case "subsup": return `<m:sSubSup>${E(node.base)}<m:sub>${R(node.sub)}</m:sub><m:sup>${R(node.sup)}</m:sup></m:sSubSup>`;
      case "nary": {
        const isInt = /[∫∬∭∮]/.test(node.chr);
        const pr = `<m:naryPr><m:chr m:val="${node.chr}"/><m:limLoc m:val="${isInt ? "subSup" : "undOvr"}"/>${node.sub ? "" : '<m:subHide m:val="1"/>'}${node.sup ? "" : '<m:supHide m:val="1"/>'}</m:naryPr>`;
        return `<m:nary>${pr}<m:sub>${R(node.sub)}</m:sub><m:sup>${R(node.sup)}</m:sup>${E(node.e || { t: "row", items: [] })}</m:nary>`;
      }
      case "limlow": return `<m:func><m:fName><m:limLow>${E(node.fn)}<m:lim>${R(node.lim)}</m:lim></m:limLow></m:fName>${E(node.arg || { t: "row", items: [] })}</m:func>`;
      case "eqarr": return `<m:eqArr>${node.rows.map(E).join("")}</m:eqArr>`;
      case "matrix": {
        const w = Math.max(1, ...node.rows.map((r) => r.length));
        return `<m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="${w}"/><m:mcJc m:val="center"/></m:mcPr></m:mc></m:mcs></m:mPr>${node.rows.map((r) => `<m:mr>${Array.from({ length: w }, (_, i) => E(r[i] || { t: "row", items: [] })).join("")}</m:mr>`).join("")}</m:m>`;
      }
      case "omml": return node.xml;
      case "aln": o._st.aln = true; return "";
      case "delim": return `<m:d><m:dPr><m:begChr m:val="${xe(node.open)}"/><m:endChr m:val="${xe(node.close)}"/></m:dPr>${E(node.e)}</m:d>`;
      case "acc": return node.chr === "\u0305" && node.name === "overline"
        ? `<m:bar><m:barPr><m:pos m:val="top"/></m:barPr>${E(node.e)}</m:bar>`
        : `<m:acc><m:accPr><m:chr m:val="${node.chr}"/></m:accPr>${E(node.e)}</m:acc>`;
      default: return "";
    }
  }

  /* ---------------- keluaran: MathML (pratinjau) ---------------- */
  function toMathML(node) {
    const M = toMathML;
    if (!node) return "<mrow></mrow>";
    switch (node.t) {
      case "row": return `<mrow>${node.items.map(M).join("")}</mrow>`;
      case "id": return node.styled ? `<mi mathvariant="normal">${xe(node.v)}</mi>` : `<mi>${xe(node.v)}</mi>`;
      case "num": return `<mn>${xe(node.v)}</mn>`;
      case "op": return `<mo>${xe(node.v)}</mo>`;
      case "space": return `<mspace width="0.3em"></mspace>`;
      case "text": return `<mtext>${xe(node.v)}</mtext>`;
      case "fn": return `<mi mathvariant="normal">${xe(node.v)}</mi>${node.arg ? `<mo>&#x2061;</mo>${M(node.arg)}` : ""}`;
      case "style": {
        // MathML Core (Chromium/WebView2) hanya mengenal mathvariant="normal": huruf ℝ/𝒟/𝐖 ditulis langsung
        const map = (x) => (x && (x.t === "id" || x.t === "num") ? Object.assign({}, x, { v: [...x.v].map((c) => styledChar(c, node.scr || "bold")).join(""), styled: true })
          : x && x.t === "row" ? Object.assign({}, x, { items: x.items.map(map) }) : x && x.t === "sub" ? Object.assign({}, x, { base: map(x.base) }) : x);
        return M(map(node.e));
      }
      case "frac": return `<mfrac${node.noBar ? ' linethickness="0"' : ""}>${M(node.n)}${M(node.d)}</mfrac>`;
      case "sqrt": return node.deg ? `<mroot>${M(node.e)}${M(node.deg)}</mroot>` : `<msqrt>${M(node.e)}</msqrt>`;
      case "sup": return `<msup>${M(node.base)}${M(node.sup)}</msup>`;
      case "sub": return `<msub>${M(node.base)}${M(node.sub)}</msub>`;
      case "subsup": return `<msubsup>${M(node.base)}${M(node.sub)}${M(node.sup)}</msubsup>`;
      case "nary": {
        const op = `<mo largeop="true" movablelimits="false">${node.chr}</mo>`;
        const s = node.sub ? M(node.sub) : null, p = node.sup ? M(node.sup) : null;
        const head = s && p ? `<munderover>${op}${s}${p}</munderover>` : s ? `<munder>${op}${s}</munder>` : p ? `<mover>${op}${p}</mover>` : op;
        return head + M(node.e);
      }
      case "limlow": return `<munder>${M(node.fn)}${M(node.lim)}</munder>${node.arg ? M(node.arg) : ""}`;
      case "eqarr": return `<mtable columnalign="left">${node.rows.map((r) => `<mtr><mtd>${M(r)}</mtd></mtr>`).join("")}</mtable>`;
      case "matrix": return `<mtable>${node.rows.map((r) => `<mtr>${r.map((c) => `<mtd>${M(c)}</mtd>`).join("")}</mtr>`).join("")}</mtable>`;
      case "delim": return `<mrow>${node.open ? `<mo stretchy="true">${xe(node.open)}</mo>` : ""}${M(node.e)}${node.close ? `<mo stretchy="true">${xe(node.close)}</mo>` : ""}</mrow>`;
      case "acc": return `<mover accent="true">${M(node.e)}<mo>${{ "\u0302": "^", "\u0305": "¯", "\u0303": "~", "\u0307": "˙", "\u0308": "¨", "\u20D7": "→", "\u030C": "ˇ", "\u0301": "´", "\u0300": "`", "\u0306": "˘" }[node.chr] || "^"}</mo></mover>`;
      default: return "";
    }
  }

  /* ---------------- keluaran: teks Unicode (untuk rumus di dalam kalimat) ---------------- */
  function toUnicode(node) {
    const U = toUnicode;
    if (!node) return "";
    const scr = (n, map, mark) => { const s = U(n); return [...s].every((c) => map[c]) ? [...s].map((c) => map[c]).join("") : mark + (s.length > 1 ? "(" + s + ")" : s); };
    switch (node.t) {
      case "row": return node.items.map(U).join("").replace(/\s+/g, " ");
      case "id": case "num": case "text": return node.v;
      case "op": return /^[=+−±×÷≤≥≠≈<>→]$/.test(node.v) ? " " + node.v + " " : node.v;
      case "space": return " ";
      case "fn": return node.v + (node.arg ? " " + U(node.arg) : "");
      case "style": return node.scr ? [...U(node.e)].map((c) => styledChar(c, node.scr)).join("") : U(node.e);
      case "frac": { const a = U(node.n), b = U(node.d); return (a.length > 1 ? "(" + a + ")" : a) + "/" + (b.length > 1 ? "(" + b + ")" : b); }
      case "sqrt": return "√" + (U(node.e).length > 1 ? "(" + U(node.e) + ")" : U(node.e));
      case "sup": return U(node.base) + scr(node.sup, TO_SUP, "^");
      case "sub": return U(node.base) + scr(node.sub, TO_SUB, "_");
      case "subsup": return U(node.base) + scr(node.sub, TO_SUB, "_") + scr(node.sup, TO_SUP, "^");
      case "nary": return node.chr + (node.sub ? scr(node.sub, TO_SUB, "_") : "") + (node.sup ? scr(node.sup, TO_SUP, "^") : "") + " " + U(node.e);
      case "limlow": return U(node.fn) + scr(node.lim, TO_SUB, "_") + " " + (node.arg ? U(node.arg) : "");
      case "eqarr": return node.rows.map(U).join("; ");
      case "matrix": return node.rows.map((r) => r.map(U).join(", ")).join("; ");
      case "omml": return node.text || "";
      case "delim": return node.open + U(node.e) + node.close;
      case "acc": return U(node.e) + node.chr;
      default: return "";
    }
  }

  /* ---------------- keluaran: runs berformat (rumus di dalam kalimat) ---------------- */
  /**
   * AST -> runs [{text, i, b, sub, sup}] untuk ditulis sebagai teks biasa di paragraf:
   * huruf variabel miring, subskrip/superskrip asli Word, ℝ/𝒟 sebagai huruf Unicode.
   */
  function toRuns(node) {
    const out = [];
    const push = (text, f) => { if (!text) return; const last = out[out.length - 1]; const r = Object.assign({ text, i: false, b: false, sub: false, sup: false }, f); if (last && last.i === r.i && last.b === r.b && last.sub === r.sub && last.sup === r.sup) last.text += text; else out.push(r); };
    const REL = /^[=≠<>≤≥≈≡∼≃≅∝∈∉⊂⊆⊃⊇→←⇒⇔↦+−±∓×÷⋅⊙⊕⊗∪∩]$/;
    const plainLen = (n) => toUnicode(n).trim().length;
    function go(n, f) {
      if (!n) return;
      switch (n.t) {
        case "row": n.items.forEach((x) => go(x, f)); return;
        case "id": for (const ch of [...n.v]) push(ch, Object.assign({}, f, { i: f.upright ? false : /[A-Za-zα-ω]/.test(ch) && !f.b ? true : f.i })); return;
        case "num": push(n.v, f); return;
        case "op": push(!f.sub && !f.sup && REL.test(n.v) ? " " + n.v + " " : n.v === "," && !f.sub && !f.sup ? ", " : n.v, Object.assign({}, f, { i: false })); return;
        case "space": push(" ", f); return;
        case "text": push(n.v, Object.assign({}, f, { i: false })); return;
        case "fn": push(n.v, Object.assign({}, f, { i: false })); if (n.arg) { if (n.arg.t !== "delim") push(" ", f); go(n.arg, f); } return;
        case "style":
          if (n.scr) { push([...toUnicode(n.e)].map((c) => styledChar(c, n.scr)).join(""), Object.assign({}, f, { i: false })); return; }
          go(n.e, Object.assign({}, f, { b: true, i: false, upright: true })); return;
        case "frac": {
          const wrap = (x) => plainLen(x) > 1 && x.t !== "delim";
          if (wrap(n.n)) push("(", f); go(n.n, f); if (wrap(n.n)) push(")", f);
          push("/", Object.assign({}, f, { i: false }));
          if (wrap(n.d)) push("(", f); go(n.d, f); if (wrap(n.d)) push(")", f);
          return;
        }
        case "sqrt": push("√", f); { const w = plainLen(n.e) > 1; if (w) push("(", f); go(n.e, f); if (w) push(")", f); } return;
        case "sup": go(n.base, f); go(n.sup, Object.assign({}, f, { sup: true, sub: false })); return;
        case "sub": go(n.base, f); go(n.sub, Object.assign({}, f, { sub: true, sup: false })); return;
        case "subsup": go(n.base, f); go(n.sub, Object.assign({}, f, { sub: true, sup: false })); go(n.sup, Object.assign({}, f, { sup: true, sub: false })); return;
        case "nary": push(n.chr, Object.assign({}, f, { i: false })); if (n.sub) go(n.sub, Object.assign({}, f, { sub: true })); if (n.sup) go(n.sup, Object.assign({}, f, { sup: true })); push(" ", f); go(n.e, f); return;
        case "limlow": push(n.fn.v || "lim", Object.assign({}, f, { i: false })); go(n.lim, Object.assign({}, f, { sub: true })); push(" ", f); if (n.arg) go(n.arg, f); return;
        case "delim": push(n.open, Object.assign({}, f, { i: false })); go(n.e, f); push(n.close, Object.assign({}, f, { i: false })); return;
        case "acc": { const before = out.length; go(n.e, f); const last = out[out.length - 1]; if (out.length > before && last) last.text += n.chr; return; }
        case "eqarr": n.rows.forEach((r, k) => { if (k) push("; ", f); go(r, f); }); return;
        case "matrix": n.rows.forEach((r, k) => { if (k) push("; ", f); r.forEach((c, j) => { if (j) push(", ", f); go(c, f); }); }); return;
        case "omml": push(n.text || "", f); return;
        default: return;
      }
    }
    go(node, {});
    // rapikan spasi ganda
    out.forEach((r, k) => { r.text = r.text.replace(/ {2,}/g, " "); if (k === 0) r.text = r.text.replace(/^ /, ""); if (k === out.length - 1) r.text = r.text.replace(/ $/, ""); });
    return out.filter((r) => r.text);
  }

  /* ---------------- paket OOXML untuk Word ---------------- */
  const PKG_HEAD = '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" mc:Ignorable="w14 w15"><w:body>';
  const PKG_TAIL = "</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>";
  const wrapPackage = (paras) => PKG_HEAD + paras + PKG_TAIL;

  /** Isi <m:oMath> dari AST atau OMML jadi (node {t:'omml'}). */
  const mathBody = (eq, o) => (eq.omml != null ? eq.omml : toOMML(eq.ast, o));

  /**
   * Satu paragraf persamaan siap sisip.
   *   eq   : { ast } atau { omml }  (+ number)
   *   opts : { mode: 'display'|'inline', numbering: 'tab'|'after'|'native'|'none', number,
   *            textWidthCm, align: 'Centered'|'Left', font, size, before, after, color }
   */
  function buildEquationParagraph(eq, opts) {
    opts = opts || {};
    const size = opts.size || 11, half = Math.round(size * 2);
    const o = { size };
    const body = mathBody(eq, o);
    if (opts.mode === "inline") return `<w:p><m:oMath>${body}</m:oMath></w:p>`;
    const number = opts.number != null && String(opts.number).trim() !== "" ? String(opts.number).trim().replace(/^\((.*)\)$/, "$1") : "";
    const numbering = number ? (opts.numbering || "tab") : "none";
    const font = xe(opts.font || "Times New Roman");
    const rpr = `<w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="${font}"/>${opts.color ? `<w:color w:val="${String(opts.color).replace("#", "")}"/>` : ""}<w:sz w:val="${half}"/><w:szCs w:val="${half}"/></w:rPr>`;
    const tw = (pt) => Math.round((pt || 0) * 20);
    const spacing = `<w:spacing w:before="${tw(opts.before != null ? opts.before : 6)}" w:after="${tw(opts.after != null ? opts.after : 6)}" w:line="${Math.round(240 * (opts.line || 1))}" w:lineRule="auto"/><w:ind w:left="0" w:right="0" w:firstLine="0"/>`;
    const jcM = opts.align === "Left" ? "left" : "center";
    const jcW = opts.align === "Left" ? "left" : "center";
    if (numbering === "tab") {
      const W = Math.round((opts.textWidthCm || 16) * 567);
      const mid = opts.align === "Left" ? Math.round(W * 0.08) : Math.round(W / 2);
      const midTab = opts.align === "Left" ? `<w:tab w:val="left" w:pos="${mid}"/>` : `<w:tab w:val="center" w:pos="${mid}"/>`;
      return `<w:p><w:pPr><w:keepLines/><w:tabs>${midTab}<w:tab w:val="right" w:pos="${W}"/></w:tabs>${spacing}<w:jc w:val="left"/>${rpr}</w:pPr><w:r>${rpr}<w:tab/></w:r><m:oMath>${body}</m:oMath><w:r>${rpr}<w:tab/><w:t>(${xe(number)})</w:t></w:r></w:p>`;
    }
    if (numbering === "after") {
      return `<w:p><w:pPr><w:keepLines/>${spacing}<w:jc w:val="${jcW}"/>${rpr}</w:pPr><m:oMath>${body}</m:oMath><w:r>${rpr}<w:t xml:space="preserve">  (${xe(number)})</w:t></w:r></w:p>`;
    }
    if (numbering === "native") { // penomoran bawaan Word 365: "#(n)" di dalam eqArr
      const num = `${ommlRun("#", o)}<m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e>${ommlRun(number, o)}</m:e></m:d>`;
      return `<w:p><w:pPr><w:keepLines/>${spacing}<w:jc w:val="${jcW}"/></w:pPr><m:oMathPara><m:oMathParaPr><m:jc m:val="${jcM}"/></m:oMathParaPr><m:oMath><m:eqArr><m:eqArrPr><m:maxDist m:val="1"/></m:eqArrPr><m:e>${body}${num}</m:e></m:eqArr></m:oMath></m:oMathPara></w:p>`;
    }
    return `<w:p><w:pPr><w:keepLines/>${spacing}<w:jc w:val="${jcW}"/></w:pPr><m:oMathPara><m:oMathParaPr><m:jc m:val="${jcM}"/></m:oMathParaPr><m:oMath>${body}</m:oMath></m:oMathPara></w:p>`;
  }
  const buildEquationPackage = (eq, opts) => wrapPackage(buildEquationParagraph(eq, opts));
  /** (kompatibel) */
  function buildPackage(ast, opts) { return buildEquationPackage({ ast }, Object.assign({ numbering: "native" }, opts)); }

  /* ---------------- OMML dari clipboard HTML Word ---------------- */
  const ENT = { nbsp: " ", amp: "&", lt: "<", gt: ">", quot: '"', apos: "'", minus: "−", times: "×", middot: "·", plusmn: "±", ne: "≠", le: "≤", ge: "≥" };
  const decodeEnt = (t) => t.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (m, e) => (e[0] === "#" ? String.fromCodePoint(e[1] === "x" || e[1] === "X" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10)) : ENT[e.toLowerCase()] != null ? ENT[e.toLowerCase()] : m));

  /**
   * Isi komentar <!--[if gte msEquation 12]> ... <![endif]--> (HTML clipboard Word) -> OMML bersih.
   * return { xml (isi m:oMath), display, ok } — xml siap ditaruh di dalam <m:oMath>.
   */
  function ommlFromWordHtml(fragment) {
    let x = String(fragment || "");
    const display = /<m:oMathPara\b/.test(x);
    x = x.replace(/<\/?(?!m:)[a-zA-Z][^>]*>/g, ""); // span/i/b dst. dibuang
    x = x.replace(/(<m:[A-Za-z]+)([^>]*?)(\/?)>/g, (all, tag, attrs, sl) => tag + attrs.replace(/\s+/g, " ").replace(/(\s[\w:]+)=([^"'\s>\/]+)/g, '$1="$2"').replace(/(\s[\w:]+)='([^']*)'/g, '$1="$2"') + sl + ">");
    x = x.replace(/<m:r>([\s\S]*?)<\/m:r>/g, (all, inner) => {
      const rp = /^\s*(<m:rPr>[\s\S]*?<\/m:rPr>)/.exec(inner);
      let rest = rp ? inner.slice(rp[0].length) : inner;
      if (/<m:t[\s>]/.test(rest)) return all;
      rest = decodeEnt(rest.replace(/\r?\n/g, " "));
      return `<m:r>${rp ? rp[1] : ""}<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t xml:space="preserve">${xe(rest)}</m:t></m:r>`;
    });
    x = x.replace(/(<\/?m:(?!t[\s>])[^>]*>)\s+(?=<)/g, "$1").replace(/^\s+|\s+$/g, "");
    // ambil isi oMath (gabung bila lebih dari satu)
    const parts = [];
    x.replace(/<m:oMath>([\s\S]*?)<\/m:oMath>/g, (m, inner) => { parts.push(inner); return m; });
    const xml = parts.length ? parts.join("") : x.replace(/<\/?m:oMath(Para)?(Pr)?>/g, "");
    return { xml, display, ok: !!xml.trim() };
  }

  /* ---------------- OMML (DOM) -> AST ---------------- */
  /** Teks run OMML ("Y=β0+1") -> token (id/num/op) agar spasi & LaTeX rapi. */
  function tokRow(t) {
    const items = [];
    const re = /(\d+(?:[.,]\d+)?)|([\p{L}\p{M}]+)|(\s+)|(.)/gu;
    let m;
    while ((m = re.exec(t))) {
      if (m[1]) items.push({ t: "num", v: m[1] });
      else if (m[2]) { for (const ch of m[2].match(/\p{L}\p{M}*/gu) || []) items.push({ t: "id", v: ch }); }
      else if (m[3]) items.push({ t: "space", v: " " });
      else items.push({ t: "op", v: m[4] === "-" ? "−" : m[4] });
    }
    return items.length === 1 ? items[0] : { t: "row", items };
  }
  function ommlToAst(el) {
    const ln = (e) => (e.localName || e.nodeName || "").replace(/^m:/, "");
    const kids = (e) => Array.from(e.childNodes || []).filter((c) => c.nodeType === 1);
    const child = (e, name) => kids(e).find((c) => ln(c) === name);
    const val = (e, name) => { const c = e && child(e, name); return c ? (c.getAttribute("m:val") != null ? c.getAttribute("m:val") : c.getAttributeNS ? c.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/math", "val") : null) : null; };
    const row = (e) => ({ t: "row", items: e ? kids(e).map(conv).filter(Boolean) : [] });
    function conv(e) {
      const n = ln(e);
      switch (n) {
        case "oMathPara": case "oMath": case "e": case "num": case "den": case "sub": case "sup": case "deg": case "lim": case "fName": return row(e);
        case "r": {
          const t = kids(e).filter((c) => ln(c) === "t").map((c) => c.textContent).join("");
          const rPr = child(e, "rPr");
          if (rPr && (child(rPr, "nor") || val(rPr, "sty") === "p") && /[A-Za-z]{2,}/.test(t)) return /^[A-Za-z]+$/.test(t.trim()) && FUNCS.has(t.trim()) ? { t: "fn", v: t.trim() } : { t: "text", v: t };
          return tokRow(t);
        }
        case "f": return { t: "frac", n: row(child(e, "num")), d: row(child(e, "den")), noBar: val(child(e, "fPr"), "type") === "noBar" };
        case "sSup": return { t: "sup", base: row(child(e, "e")), sup: row(child(e, "sup")) };
        case "sSub": return { t: "sub", base: row(child(e, "e")), sub: row(child(e, "sub")) };
        case "sSubSup": return { t: "subsup", base: row(child(e, "e")), sub: row(child(e, "sub")), sup: row(child(e, "sup")) };
        case "sPre": return { t: "row", items: [{ t: "subsup", base: { t: "row", items: [] }, sub: row(child(e, "sub")), sup: row(child(e, "sup")) }, row(child(e, "e"))] };
        case "rad": { const pr = child(e, "radPr"), deg = child(e, "deg"); const hide = pr && /^(1|on|true)$/.test(val(pr, "degHide") || ""); return { t: "sqrt", e: row(child(e, "e")), deg: !hide && deg && kids(deg).length ? row(deg) : null }; }
        case "nary": {
          const pr = child(e, "naryPr");
          const chr = (pr && val(pr, "chr")) || "∫";
          const sub = child(e, "sub"), sup = child(e, "sup");
          return { t: "nary", chr, sub: sub && kids(sub).length ? row(sub) : null, sup: sup && kids(sup).length ? row(sup) : null, e: row(child(e, "e")) };
        }
        case "d": {
          const pr = child(e, "dPr");
          const beg = pr && val(pr, "begChr"), end = pr && val(pr, "endChr");
          const es = kids(e).filter((c) => ln(c) === "e").map(row);
          const sep = (pr && val(pr, "sepChr")) || ",";
          const items = es.reduce((a, r, i) => a.concat(i ? [{ t: "op", v: sep }] : [], [r]), []);
          return { t: "delim", open: beg != null ? beg : "(", close: end != null ? end : ")", e: { t: "row", items } };
        }
        case "acc": { const pr = child(e, "accPr"); return { t: "acc", chr: (pr && val(pr, "chr")) || "\u0302", e: row(child(e, "e")) }; }
        case "bar": return { t: "acc", chr: "\u0305", name: "overline", e: row(child(e, "e")) };
        case "func": {
          const fn = row(child(e, "fName")), arg = row(child(e, "e"));
          const only = fn.items.length === 1 ? fn.items[0] : null;
          if (only && only.t === "limlow") return Object.assign({}, only, { arg });
          const name = only && (only.t === "id" || only.t === "text" || only.t === "fn") ? only.v : null;
          return name ? { t: "fn", v: name, arg } : { t: "row", items: [fn, arg] };
        }
        case "limLow": { const b = row(child(e, "e")); const nm = b.items.length === 1 && b.items[0].v; return { t: "limlow", fn: { t: "fn", v: nm || "lim" }, lim: row(child(e, "lim")) }; }
        case "limUpp": return { t: "sup", base: row(child(e, "e")), sup: row(child(e, "lim")) };
        case "groupChr": case "box": case "borderBox": case "phant": return row(child(e, "e"));
        case "eqArr": return { t: "eqarr", rows: kids(e).filter((c) => ln(c) === "e").map(row) };
        case "m": return { t: "matrix", rows: kids(e).filter((c) => ln(c) === "mr").map((mr) => kids(mr).filter((c) => ln(c) === "e").map(row)) };
        default: return null; // *Pr, ctrlPr, dll.
      }
    }
    return conv(el) || { t: "row", items: [] };
  }

  /** Teks OMML (string isi oMath) -> AST, memakai DOMParser (browser). */
  function ommlXmlToAst(xml) {
    if (typeof DOMParser === "undefined") return null;
    const doc = new DOMParser().parseFromString(`<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${xml}</m:oMath>`, "application/xml");
    if (doc.getElementsByTagName("parsererror").length) return null;
    return ommlToAst(doc.documentElement);
  }

  /* ---------------- AST -> LaTeX (untuk tombol Salin LaTeX & Ambil rumus) ---------------- */
  const REV = {};
  Object.entries(GREEK).forEach(([k, v]) => { if (!REV[v]) REV[v] = "\\" + k; });
  REV["ε"] = "\\varepsilon"; REV["ϵ"] = "\\epsilon"; REV["φ"] = "\\varphi"; REV["ϕ"] = "\\phi";
  Object.entries(OPS).forEach(([k, v]) => { if (!REV[v] && /^[a-zA-Z]+$/.test(k) && v.length === 1 && !/[{}%$&#_|]/.test(v)) REV[v] = "\\" + k; });
  REV["−"] = "-"; REV["⋅"] = "\\cdot"; REV["∗"] = "*"; REV["′"] = "'";
  const ACC_REV = { "\u0302": "hat", "\u0305": "bar", "\u0303": "tilde", "\u0307": "dot", "\u0308": "ddot", "\u20D7": "vec", "\u030C": "check", "\u0301": "acute", "\u0300": "grave", "\u0306": "breve" };
  const DELIM_REV = { "(": "(", ")": ")", "[": "[", "]": "]", "{": "\\{", "}": "\\}", "|": "|", "‖": "\\|", "⟨": "\\langle", "⟩": "\\rangle", "⌊": "\\lfloor", "⌋": "\\rfloor", "⌈": "\\lceil", "⌉": "\\rceil", "": "." };
  const NARY_REV = Object.fromEntries(Object.entries(NARY).map(([k, v]) => [v, "\\" + k]));
  function toLatex(node) {
    const L = toLatex;
    const g = (n) => { const s = L(n).trim(); return s.length === 1 || /^\\[a-zA-Z]+(\{[^{}]*\})?$/.test(s) ? s : "{" + s + "}"; };
    const chars = (s) => [...String(s)].map((c) => (REV[c] ? REV[c] + (/[a-z]$/i.test(REV[c]) ? " " : "") : c === "%" || c === "&" || c === "#" ? "\\" + c : c)).join("");
    if (!node) return "";
    switch (node.t) {
      case "row": return node.items.map(L).join("").replace(/ +/g, " ").replace(/ $/, "");
      case "id": case "num": return chars(node.v);
      case "op": return chars(node.v);
      case "space": return node.v ? "\\," : "";
      case "text": return "\\text{" + node.v + "}";
      case "fn": { const nm = TEX_FUNCS.has(node.v) ? "\\" + node.v : "\\operatorname{" + node.v + "}"; return nm + (node.arg ? (node.arg.t === "delim" ? "" : " ") + L(node.arg) : " "); }
      case "style": { const inv = { "double-struck": "mathbb", script: "mathcal", fraktur: "mathfrak", "sans-serif": "mathsf", monospace: "mathtt" }; return "\\" + (node.scr ? inv[node.scr] || "mathrm" : "mathbf") + "{" + L(node.e) + "}"; }
      case "frac": return node.noBar ? "\\binom" + g(node.n) + g(node.d) : "\\frac{" + L(node.n) + "}{" + L(node.d) + "}";
      case "sqrt": return "\\sqrt" + (node.deg ? "[" + L(node.deg) + "]" : "") + "{" + L(node.e) + "}";
      case "sup": return g(node.base) + "^" + g(node.sup);
      case "sub": return g(node.base) + "_" + g(node.sub);
      case "subsup": return g(node.base) + "_" + g(node.sub) + "^" + g(node.sup);
      case "nary": return (NARY_REV[node.chr] || node.chr) + (node.sub ? "_" + g(node.sub) : "") + (node.sup ? "^" + g(node.sup) : "") + " " + L(node.e);
      case "limlow": return "\\" + (TEX_FUNCS.has(node.fn.v) ? node.fn.v : "lim") + "_" + g(node.lim) + " " + (node.arg ? L(node.arg) : "");
      case "delim": {
        let inner = node.e;
        while (inner && inner.t === "row" && inner.items.length === 1) inner = inner.items[0];
        if (inner && inner.t === "eqarr" && node.open === "{" && !node.close) return "\\begin{cases}" + inner.rows.map(L).join(" \\\\ ") + "\\end{cases}";
        if (node.e && node.e.t === "eqarr" && node.open === "{" && !node.close) return "\\begin{cases}" + node.e.rows.map(L).join(" \\\\ ") + "\\end{cases}";
        if (node.e && node.e.t === "frac" && node.e.noBar) return L(node.e);
        return "\\left" + (DELIM_REV[node.open] != null ? DELIM_REV[node.open] : node.open) + L(node.e) + "\\right" + (DELIM_REV[node.close] != null ? DELIM_REV[node.close] : node.close);
      }
      case "acc": return "\\" + (node.name === "overline" ? "overline" : ACC_REV[node.chr] || "hat") + "{" + L(node.e) + "}";
      case "eqarr": return "\\begin{aligned}" + node.rows.map(L).join(" \\\\ ") + "\\end{aligned}";
      case "aln": return "&";
      case "matrix": return "\\begin{matrix}" + node.rows.map((r) => r.map(L).join(" & ")).join(" \\\\ ") + "\\end{matrix}";
      case "omml": return node.text || "";
      default: return "";
    }
  }

  /** Dari teks masukan apa pun -> { ast, tex, number, error } */
  function parse(raw, opts) {
    try {
      const n = normalizeInput(raw, opts);
      if (!n.tex) return { error: "Rumus kosong." };
      return { ast: parseLatex(n.tex), tex: n.tex, number: n.number };
    } catch (e) { return { error: "Rumus tidak bisa dibaca: " + e.message }; }
  }

  /* ---------------- pustaka rumus penelitian ---------------- */
  const LIBRARY = [
    { group: "Statistik dasar", name: "Rata-rata", tex: "\\bar{x}=\\frac{1}{n}\\sum_{i=1}^{n} x_i" },
    { group: "Statistik dasar", name: "Simpangan baku (sampel)", tex: "s=\\sqrt{\\frac{\\sum_{i=1}^{n}\\left(x_i-\\bar{x}\\right)^2}{n-1}}" },
    { group: "Statistik dasar", name: "Z-score", tex: "Z=\\frac{X-\\mu}{\\sigma}" },
    { group: "Statistik dasar", name: "Persentase", tex: "P=\\frac{f}{N}\\times 100\\%" },
    { group: "Sampel", name: "Slovin", tex: "n=\\frac{N}{1+N e^2}" },
    { group: "Sampel", name: "Lemeshow", tex: "n=\\frac{Z_{1-\\alpha/2}^2\\, p\\left(1-p\\right)}{d^2}" },
    { group: "Instrumen", name: "Cronbach's alpha", tex: "\\alpha=\\frac{k}{k-1}\\left(1-\\frac{\\sum_{i=1}^{k}\\sigma_{i}^2}{\\sigma_{t}^2}\\right)" },
    { group: "Instrumen", name: "Korelasi Pearson", tex: "r_{xy}=\\frac{n\\sum xy-\\left(\\sum x\\right)\\left(\\sum y\\right)}{\\sqrt{\\left[n\\sum x^2-\\left(\\sum x\\right)^2\\right]\\left[n\\sum y^2-\\left(\\sum y\\right)^2\\right]}}" },
    { group: "Instrumen", name: "AVE (SEM-PLS)", tex: "AVE=\\frac{\\sum_{i=1}^{n}\\lambda_i^2}{n}" },
    { group: "Instrumen", name: "Composite reliability", tex: "CR=\\frac{\\left(\\sum\\lambda_i\\right)^2}{\\left(\\sum\\lambda_i\\right)^2+\\sum\\left(1-\\lambda_i^2\\right)}" },
    { group: "Uji", name: "Uji t (dua sampel)", tex: "t=\\frac{\\bar{x}_1-\\bar{x}_2}{\\sqrt{\\frac{s_1^2}{n_1}+\\frac{s_2^2}{n_2}}}" },
    { group: "Uji", name: "Chi-square", tex: "\\chi^2=\\sum\\frac{\\left(O_i-E_i\\right)^2}{E_i}" },
    { group: "Uji", name: "Cohen's d", tex: "d=\\frac{\\bar{x}_1-\\bar{x}_2}{s_{pooled}}" },
    { group: "Uji", name: "N-Gain", tex: "g=\\frac{S_{post}-S_{pre}}{S_{max}-S_{pre}}" },
    { group: "Regresi", name: "Regresi linear berganda", tex: "Y=\\beta_0+\\beta_1X_1+\\beta_2X_2+\\varepsilon" },
    { group: "Regresi", name: "Regresi moderasi (MRA)", tex: "Y=\\beta_0+\\beta_1X+\\beta_2Z+\\beta_3\\left(X\\times Z\\right)+\\varepsilon" },
    { group: "Regresi", name: "Koefisien determinasi", tex: "R^2=1-\\frac{SS_{res}}{SS_{tot}}" },
    { group: "Regresi", name: "Regresi logistik", tex: "\\ln\\left(\\frac{p}{1-p}\\right)=\\beta_0+\\beta_1X_1+\\cdots+\\beta_kX_k" },
    { group: "Klaster", name: "Fuzzy C-Means (fungsi objektif)", tex: "J_m=\\sum_{i=1}^{N}\\sum_{k=1}^{c}u_{ik}^{m}\\left\\|x_i-v_k\\right\\|^2" },
    { group: "Klaster", name: "Jarak Euclidean", tex: "d\\left(x,y\\right)=\\sqrt{\\sum_{i=1}^{n}\\left(x_i-y_i\\right)^2}" },
  ];

  return { mathAlnumToPlain, toRuns, repairLatex, parse, parseLatex, normalizeInput, linearToLatex, fromMathML, toOMML, toMathML, toUnicode, toLatex, buildPackage, buildEquationParagraph, buildEquationPackage, wrapPackage, ommlFromWordHtml, ommlToAst, ommlXmlToAst, LIBRARY, looksLatex };
});
