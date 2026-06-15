/**
 * Font utilities for the statistics report exports.
 *
 * Statistics PDFs and Word documents mix Latin and Georgian content
 * ("USA-დან საქართველოში 2024 წელს ექსპორტი შეადგინა $123 მლნ").
 * Arial covers Latin/digits/punctuation universally but has zero
 * Georgian glyphs; Sylfaen covers Georgian (Mkhedruli + Mtavruli) but
 * looks slightly off for Latin numerics. The standard Microsoft
 * typography pattern pairs Arial for Latin and Sylfaen for Georgian.
 *
 * splitByScript() walks a string and returns an ordered list of
 * { text, font } segments so the export builders can emit one run per
 * script and let each font cover what it's good at.
 */
(function () {
  'use strict';

  // Mkhedruli (U+10A0–U+10FF) + Mtavruli (U+2D00–U+2D2F) +
  // Extended Mkhedruli for Mtavruli case-folding (U+1C90–U+1CBF).
  const GEO_RE = /[Ⴀ-ჿⴀ-⴯Ა-Ჿ]+/g;

  function splitByScript(text, opts) {
    const latin = (opts && opts.latin) || 'Arial';
    const georgian = (opts && opts.georgian) || 'Sylfaen';
    const s = text == null ? '' : String(text);
    if (!s) return [{ text: '', font: latin }];
    const out = [];
    let i = 0;
    let m;
    GEO_RE.lastIndex = 0;
    while ((m = GEO_RE.exec(s)) !== null) {
      if (m.index > i) out.push({ text: s.slice(i, m.index), font: latin });
      out.push({ text: m[0], font: georgian });
      i = GEO_RE.lastIndex;
    }
    if (i < s.length) out.push({ text: s.slice(i), font: latin });
    return out;
  }

  function hasGeorgian(text) {
    if (text == null) return false;
    GEO_RE.lastIndex = 0;
    return GEO_RE.test(String(text));
  }

  window.FontUtils = { splitByScript, hasGeorgian };
})();
