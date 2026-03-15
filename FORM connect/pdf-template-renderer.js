(function initCsd1PdfTemplateRenderer(global) {
  'use strict';

  const PDFLibNS = global.PDFLib;
  const fontkitNS = global.fontkit || global.Fontkit;

  const DEFAULT_SPEC = {
    page: { width: 595, height: 842 },
    sharedFields: {
      doc_number_line: { x: 73, y: 90, w: 150, h: 24, align: 'left', size: 16, maxLines: 1, clearBackground: false },
      issue_date_full: { x: 174, y: 156, w: 252, h: 20, align: 'center', size: 16, maxLines: 1 },
    },
    bodyStart: { x: 54, y: 294, w: 487 },
    labelBlock: {
      labelX: 54,
      labelWidth: 38,
      bodyX: 98,
      bodyWidth: 438,
      gapAfter: 2,
    },
    list: {
      gapBefore: 6,
      gapAfter: 8,
      startX: 74,
      numberColumnWidth: 18,
      contentWidth: 450,
      sublineIndent: 24,
      itemGap: 2,
    },
    closingBlock: {
      gapBefore: 4,
      x: 98,
      w: 438,
    },
    signatureBlock: {
      gapBefore: 4,
      x: 315,
      w: 175,
      centerX: 320,
      imageFitWidth: 118,
      imageFitHeight: 30,
      groupGap: 10,
      row1Height: 30,
      rankTopOffset: 6,
      imageTopOffset: 0,
      nameGap: 2,
      positionGap: 2,
    },
    contactBlock: {
      gapBefore: 4,
      x: 54,
      w: 487,
    },
    font: {
      size: 16,
      lineHeight: 16.2,
      paragraphIndent: 23,
      paragraphGap: 1,
    },
    flowBottomLimit: 826,
  };

  const COLOR_BLACK = { r: 0, g: 0, b: 0 };
  const COLOR_WHITE = { r: 1, g: 1, b: 1 };
  const cache = {
    arrayBuffer: new Map(),
    json: new Map(),
  };

  function assertDependencies() {
    if (!PDFLibNS || !PDFLibNS.PDFDocument || !PDFLibNS.rgb) {
      throw new Error('ไม่พบ pdf-lib');
    }
    if (!fontkitNS) {
      throw new Error('ไม่พบ fontkit');
    }
  }

  async function fetchArrayBuffer(url) {
    if (cache.arrayBuffer.has(url)) return cache.arrayBuffer.get(url);
    const response = await fetch(url, { cache: 'no-store' });
    if (!response.ok) {
      throw new Error(`โหลดไฟล์ไม่สำเร็จ (${response.status})`);
    }
    const data = await response.arrayBuffer();
    cache.arrayBuffer.set(url, data);
    return data;
  }

  async function fetchJson(url) {
    if (cache.json.has(url)) return cache.json.get(url);
    const response = await fetch(url, { cache: 'no-store' });
    if (!response.ok) {
      throw new Error(`โหลด config ไม่สำเร็จ (${response.status})`);
    }
    const data = await response.json();
    cache.json.set(url, data);
    return data;
  }

  function isObject(value) {
    return !!value && typeof value === 'object' && !Array.isArray(value);
  }

  function mergeObjects(base, extra) {
    const output = { ...(base || {}) };
    Object.entries(extra || {}).forEach(([key, value]) => {
      if (isObject(value) && isObject(output[key])) {
        output[key] = mergeObjects(output[key], value);
        return;
      }
      output[key] = value;
    });
    return output;
  }

  async function resolveSpec(config) {
    let spec = mergeObjects({}, DEFAULT_SPEC);
    if (!config || typeof config !== 'object') return spec;

    if (config.base) {
      spec = mergeObjects(spec, await fetchJson(config.base));
    }
    if (config.type) {
      spec = mergeObjects(spec, await fetchJson(config.type));
    }
    if (config.inline && typeof config.inline === 'object') {
      spec = mergeObjects(spec, config.inline);
    }
    return spec;
  }

  function topToPdfY(pageHeight, topY, boxHeight) {
    return pageHeight - topY - boxHeight;
  }

  function normalizeSpace(value) {
    return String(value || '').replace(/\s+/g, ' ').trim();
  }

  function segmentWordToken(token) {
    if (!token) return [];
    if (!(global.Intl && global.Intl.Segmenter)) return [token];
    try {
      const segmenter = new global.Intl.Segmenter('th', { granularity: 'word' });
      return Array.from(segmenter.segment(token), item => item.segment).filter(Boolean);
    } catch (_) {
      return [token];
    }
  }

  function isThaiWordToken(token) {
    return /^[\u0E00-\u0E7F]+$/u.test(String(token || ''));
  }

  function compactThaiTokens(tokens) {
    const output = [];
    let buffer = '';
    const MIN_GROUP_LENGTH = 6;
    const MAX_GROUP_LENGTH = 12;

    function flushBuffer() {
      if (!buffer) return;
      output.push(buffer);
      buffer = '';
    }

    tokens.forEach((token) => {
      if (!token) return;
      if (/^\s+$/u.test(token)) {
        flushBuffer();
        output.push(token);
        return;
      }
      if (!isThaiWordToken(token)) {
        flushBuffer();
        output.push(token);
        return;
      }

      if (!buffer) {
        buffer = token;
        return;
      }

      const candidate = buffer + token;
      if (candidate.length <= MAX_GROUP_LENGTH || buffer.length < MIN_GROUP_LENGTH) {
        buffer = candidate;
        return;
      }

      flushBuffer();
      buffer = token;
    });

    flushBuffer();
    return output;
  }

  function attachQuoteTokens(tokens) {
    const output = [];
    const OPEN_QUOTES = new Set(['"', '“', '‘']);
    const CLOSE_QUOTES = new Set(['"', '”', '’']);

    (tokens || []).forEach((token) => {
      if (!token) return;
      if (/^\s+$/u.test(token)) {
        output.push(token);
        return;
      }

      if (OPEN_QUOTES.has(token)) {
        output.push(`__OPEN_QUOTE__${token}`);
        return;
      }

      if (CLOSE_QUOTES.has(token)) {
        let attached = false;
        for (let index = output.length - 1; index >= 0; index -= 1) {
          if (/^\s+$/u.test(output[index])) continue;
          output[index] = `${output[index]}${token}`;
          attached = true;
          break;
        }
        if (!attached) {
          output.push(token);
        }
        return;
      }

      if (output.length) {
        const previousIndex = output.length - 1;
        if (
          typeof output[previousIndex] === 'string' &&
          output[previousIndex].startsWith('__OPEN_QUOTE__')
        ) {
          const openQuote = output[previousIndex].replace('__OPEN_QUOTE__', '');
          output[previousIndex] = `${openQuote}${token}`;
          return;
        }
      }

      output.push(token);
    });

    return output.map(token => String(token).replace(/^__OPEN_QUOTE__/, ''));
  }

  function tokenizeText(text) {
    const source = String(text || '');
    const parts = source.split(/(\s+)/).filter(part => part.length);
    const tokens = [];
    parts.forEach((part) => {
      if (/^\s+$/.test(part)) {
        tokens.push(part);
        return;
      }
      segmentWordToken(part).forEach(token => tokens.push(token));
    });
    return attachQuoteTokens(compactThaiTokens(tokens));
  }

  function measureText(font, size, text) {
    return font.widthOfTextAtSize(String(text || ''), size);
  }

  function breakLongToken(font, size, token, maxWidth) {
    const chars = Array.from(String(token || ''));
    const pieces = [];
    let current = '';
    chars.forEach((char) => {
      const candidate = current + char;
      if (!current || measureText(font, size, candidate) <= maxWidth) {
        current = candidate;
        return;
      }
      pieces.push(current);
      current = char;
    });
    if (current) pieces.push(current);
    return pieces;
  }

  function wrapTokens(font, size, tokens, maxWidth) {
    const lines = [];
    let current = '';

    function pushCurrent() {
      const next = current.replace(/\s+$/g, '');
      lines.push(next);
      current = '';
    }

    tokens.forEach((token) => {
      if (!token) return;
      if (!current) {
        const nextToken = token.replace(/^\s+/g, '');
        if (!nextToken) return;
        if (measureText(font, size, nextToken) <= maxWidth) {
          current = nextToken;
          return;
        }
        const pieces = breakLongToken(font, size, nextToken, maxWidth);
        pieces.slice(0, -1).forEach(piece => lines.push(piece));
        current = pieces[pieces.length - 1] || '';
        return;
      }

      const candidate = current + token;
      if (measureText(font, size, candidate) <= maxWidth) {
        current = candidate;
        return;
      }

      pushCurrent();

      const nextToken = token.replace(/^\s+/g, '');
      if (!nextToken) return;
      if (measureText(font, size, nextToken) <= maxWidth) {
        current = nextToken;
        return;
      }
      const pieces = breakLongToken(font, size, nextToken, maxWidth);
      pieces.slice(0, -1).forEach(piece => lines.push(piece));
      current = pieces[pieces.length - 1] || '';
    });

    if (current || !lines.length) {
      lines.push(current.replace(/\s+$/g, ''));
    }

    return lines;
  }

  function wrapText(font, size, text, maxWidth) {
    const rawLines = String(text == null ? '' : text).split(/\r?\n/);
    const lines = [];
    rawLines.forEach((raw) => {
      const clean = raw || '';
      const tokens = tokenizeText(clean);
      const wrapped = wrapTokens(font, size, tokens, maxWidth);
      if (!wrapped.length) {
        lines.push('');
        return;
      }
      wrapped.forEach(line => lines.push(line));
    });
    return lines.length ? lines : [''];
  }

  function wrapParagraph(font, size, text, width, firstLineIndent) {
    const rawLines = String(text == null ? '' : text).split(/\r?\n/);
    const lines = [];
    rawLines.forEach((raw, paragraphIndex) => {
      const tokens = tokenizeText(raw || '');
      let current = '';
      let currentWidth = width - firstLineIndent;
      let currentIndent = firstLineIndent;

      function pushCurrentLine() {
        lines.push({
          text: current.replace(/\s+$/g, ''),
          indent: currentIndent,
        });
        current = '';
        currentWidth = width;
        currentIndent = 0;
      }

      tokens.forEach((token) => {
        if (!token) return;
        const normalized = current ? token : token.replace(/^\s+/g, '');
        if (!normalized) return;
        const candidate = current + normalized;
        if (measureText(font, size, candidate) <= currentWidth) {
          current = candidate;
          return;
        }

        if (current) {
          pushCurrentLine();
        }

        if (measureText(font, size, normalized.replace(/^\s+/g, '')) <= currentWidth) {
          current = normalized.replace(/^\s+/g, '');
          return;
        }

        const pieces = breakLongToken(font, size, normalized.replace(/^\s+/g, ''), currentWidth);
        pieces.forEach((piece, pieceIndex) => {
          if (pieceIndex < pieces.length - 1) {
            lines.push({ text: piece, indent: currentIndent });
            currentIndent = 0;
            currentWidth = width;
            return;
          }
          current = piece;
          currentIndent = 0;
          currentWidth = width;
        });
      });

      if (current || !lines.length) {
        lines.push({
          text: current.replace(/\s+$/g, ''),
          indent: currentIndent,
        });
      }

      if (paragraphIndex < rawLines.length - 1) {
        lines.push({ text: '', indent: 0, isBlank: true });
      }
    });
    return lines;
  }

  function wrapParagraphBalanced(font, size, text, width, firstLineIndent) {
    const initialLines = wrapParagraph(font, size, text, width, firstLineIndent);
    const visibleLines = initialLines.filter(line => !line.isBlank);
    if (visibleLines.length < 2 || firstLineIndent <= 0) {
      return initialLines;
    }

    const lastLineWidth = measureText(font, size, visibleLines[visibleLines.length - 1].text || '');
    const threshold = Math.min(120, width * 0.22);
    if (lastLineWidth >= threshold) {
      return initialLines;
    }

    const noIndentLines = wrapParagraph(font, size, text, width, 0);
    const noIndentVisible = noIndentLines.filter(line => !line.isBlank);
    if (!noIndentVisible.length) {
      return initialLines;
    }
    const noIndentLastWidth = measureText(font, size, noIndentVisible[noIndentVisible.length - 1].text || '');
    return noIndentLastWidth > lastLineWidth ? noIndentLines : initialLines;
  }

  function clampLines(font, size, lines, maxWidth, maxLines) {
    if (!maxLines || lines.length <= maxLines) {
      return { lines, truncated: false };
    }
    const clipped = lines.slice(0, maxLines);
    let last = clipped[maxLines - 1] || '';
    const ellipsis = '...';
    while (last && measureText(font, size, `${last}${ellipsis}`) > maxWidth) {
      last = last.slice(0, -1);
    }
    clipped[maxLines - 1] = `${last}${ellipsis}`;
    return { lines: clipped, truncated: true };
  }

  function drawSimpleLines(page, font, size, color, lines, x, topY, width, lineHeight, align, centerX) {
    lines.forEach((line, index) => {
      const text = line == null ? '' : String(line);
      const renderedWidth = measureText(font, size, text);
      let drawX = x;
      if (align === 'center') {
        drawX = centerX != null
          ? centerX - (renderedWidth / 2)
          : x + Math.max(0, (width - renderedWidth) / 2);
      } else if (align === 'right') {
        drawX = x + Math.max(0, width - renderedWidth);
      }
      page.drawText(text, {
        x: drawX,
        y: topToPdfY(page.getHeight(), topY + (index * lineHeight), size),
        size,
        font,
        color,
      });
    });
  }

  function normalizeTextSegments(value) {
    if (!Array.isArray(value)) {
      return [{
        text: String(value == null ? '' : value),
        bold: false,
      }];
    }
    return value
      .map((entry) => {
        if (entry && typeof entry === 'object' && !Array.isArray(entry)) {
          return {
            text: String(entry.text == null ? '' : entry.text),
            bold: entry.bold === true,
          };
        }
        return {
          text: String(entry == null ? '' : entry),
          bold: false,
        };
      })
      .filter(segment => segment.text.length);
  }

  function tokenizeRichTextSegments(segments) {
    const tokens = [];
    normalizeTextSegments(segments).forEach((segment) => {
      tokenizeText(segment.text).forEach((token) => {
        tokens.push({
          text: token,
          bold: segment.bold === true,
        });
      });
    });
    return tokens;
  }

  function getSegmentFont(fonts, segment) {
    return segment && segment.bold ? fonts.bold : fonts.regular;
  }

  function measureSegmentList(fonts, size, segments) {
    return (segments || []).reduce(
      (total, segment) => total + measureText(getSegmentFont(fonts, segment), size, segment.text || ''),
      0
    );
  }

  function trimTrailingSegmentSpaces(segments) {
    const output = (segments || []).map(segment => ({ ...segment }));
    while (output.length) {
      const last = output[output.length - 1];
      last.text = String(last.text || '').replace(/\s+$/g, '');
      if (last.text) break;
      output.pop();
    }
    return output;
  }

  function mergeAdjacentSegments(segments) {
    const output = [];
    (segments || []).forEach((segment) => {
      const text = String(segment.text || '');
      if (!text) return;
      const previous = output[output.length - 1];
      if (previous && previous.bold === !!segment.bold) {
        previous.text += text;
        return;
      }
      output.push({
        text,
        bold: !!segment.bold,
      });
    });
    return output;
  }

  function breakLongRichToken(fonts, size, token, maxWidth) {
    const chars = Array.from(String(token.text || ''));
    const pieces = [];
    let current = '';
    chars.forEach((char) => {
      const candidate = current + char;
      if (!current || measureText(getSegmentFont(fonts, token), size, candidate) <= maxWidth) {
        current = candidate;
        return;
      }
      pieces.push({ text: current, bold: token.bold === true });
      current = char;
    });
    if (current) {
      pieces.push({ text: current, bold: token.bold === true });
    }
    return pieces;
  }

  function wrapRichTokens(fonts, size, tokens, maxWidth) {
    const lines = [];
    let current = [];

    function pushCurrent() {
      const merged = mergeAdjacentSegments(trimTrailingSegmentSpaces(current));
      lines.push(merged);
      current = [];
    }

    tokens.forEach((token) => {
      if (!token || !token.text) return;
      const nextText = current.length ? token.text : token.text.replace(/^\s+/g, '');
      if (!nextText) return;
      const normalizedToken = { text: nextText, bold: token.bold === true };
      const candidate = current.concat(normalizedToken);
      if (measureSegmentList(fonts, size, candidate) <= maxWidth) {
        current = candidate;
        return;
      }

      if (current.length) {
        pushCurrent();
      }

      const trimmedToken = { text: normalizedToken.text.replace(/^\s+/g, ''), bold: normalizedToken.bold };
      if (!trimmedToken.text) return;
      if (measureSegmentList(fonts, size, [trimmedToken]) <= maxWidth) {
        current = [trimmedToken];
        return;
      }

      const pieces = breakLongRichToken(fonts, size, trimmedToken, maxWidth);
      pieces.slice(0, -1).forEach(piece => lines.push([piece]));
      current = pieces.length ? [pieces[pieces.length - 1]] : [];
    });

    if (current.length || !lines.length) {
      pushCurrent();
    }

    return lines;
  }

  function wrapRichParagraph(fonts, size, segments, width, firstLineIndent) {
    const tokens = tokenizeRichTextSegments(segments);
    const lines = [];
    let current = [];
    let currentWidth = width - firstLineIndent;
    let currentIndent = firstLineIndent;

    function pushCurrentLine() {
      lines.push({
        segments: mergeAdjacentSegments(trimTrailingSegmentSpaces(current)),
        indent: currentIndent,
      });
      current = [];
      currentWidth = width;
      currentIndent = 0;
    }

    tokens.forEach((token) => {
      if (!token || !token.text) return;
      const normalizedText = current.length ? token.text : token.text.replace(/^\s+/g, '');
      if (!normalizedText) return;
      const normalizedToken = { text: normalizedText, bold: token.bold === true };
      const candidate = current.concat(normalizedToken);
      if (measureSegmentList(fonts, size, candidate) <= currentWidth) {
        current = candidate;
        return;
      }

      if (current.length) {
        pushCurrentLine();
      }

      const trimmedToken = { text: normalizedToken.text.replace(/^\s+/g, ''), bold: normalizedToken.bold };
      if (!trimmedToken.text) return;
      if (measureSegmentList(fonts, size, [trimmedToken]) <= currentWidth) {
        current = [trimmedToken];
        return;
      }

      const pieces = breakLongRichToken(fonts, size, trimmedToken, currentWidth);
      pieces.forEach((piece, pieceIndex) => {
        if (pieceIndex < pieces.length - 1) {
          lines.push({
            segments: [piece],
            indent: currentIndent,
          });
          currentIndent = 0;
          currentWidth = width;
          return;
        }
        current = [piece];
        currentIndent = 0;
        currentWidth = width;
      });
    });

    if (current.length || !lines.length) {
      lines.push({
        segments: mergeAdjacentSegments(trimTrailingSegmentSpaces(current)),
        indent: currentIndent,
      });
    }

    return lines;
  }

  function drawRichSegmentsLine(page, ctx, segments, x, topY) {
    let cursorX = x;
    (segments || []).forEach((segment) => {
      const text = String(segment.text || '');
      if (!text) return;
      const font = getSegmentFont(ctx.fonts, segment);
      page.drawText(text, {
        x: cursorX,
        y: topToPdfY(page.getHeight(), topY, ctx.metrics.bodySize),
        size: ctx.metrics.bodySize,
        font,
        color: ctx.colors.black,
      });
      cursorX += measureText(font, ctx.metrics.bodySize, text);
    });
  }

  function getLabelBlockSpec(block, ctx) {
    const base = ctx.spec.labelBlock || {};
    return {
      labelX: block.labelX != null ? block.labelX : base.labelX,
      labelWidth: block.labelWidth != null ? block.labelWidth : base.labelWidth,
      bodyX: block.bodyX != null ? block.bodyX : base.bodyX,
      bodyWidth: block.bodyWidth != null ? block.bodyWidth : base.bodyWidth,
      gapAfter: block.gapAfter != null ? block.gapAfter : base.gapAfter,
    };
  }

  function getLabelLines(block) {
    return (block.labelLines || []).filter(Boolean).map(String);
  }

  function normalizeLineEntry(line) {
    if (line && typeof line === 'object' && !Array.isArray(line)) {
      return {
        text: String(line.text || ''),
        bold: line.bold === true,
      };
    }
    return {
      text: String(line || ''),
      bold: false,
    };
  }

  function drawLabelStack(page, ctx, lines, x, topY, width) {
    drawSimpleLines(
      page,
      ctx.fonts.bold,
      ctx.metrics.bodySize,
      ctx.colors.black,
      lines,
      x,
      topY,
      width,
      ctx.metrics.lineHeight,
      'left'
    );
  }

  function measureParagraphHeight(lines, lineHeight, paragraphGap) {
    if (!lines.length) return lineHeight;
    const visibleLines = lines.filter(line => !line.isBlank).length;
    if (!visibleLines) return lineHeight;
    return (lines.length * lineHeight) + paragraphGap;
  }

  function measureFlowBlocks(blocks, ctx) {
    const spec = ctx.spec;
    let height = 0;
    blocks.forEach((block) => {
      if (block.type === 'paragraph') {
        const width = block.width != null
          ? block.width
          : ((ctx.spec.labelBlock && ctx.spec.labelBlock.bodyWidth) || ctx.spec.bodyStart.w);
        const firstLineIndent = block.firstLineIndent != null ? block.firstLineIndent : ctx.spec.font.paragraphIndent;
        const lines = wrapRichParagraph(
          ctx.fonts,
          ctx.metrics.bodySize,
          block.textSegments || block.text,
          width,
          firstLineIndent
        );
        height += measureParagraphHeight(lines, ctx.metrics.lineHeight, spec.font.paragraphGap);
      } else if (block.type === 'labelParagraph') {
        const labelSpec = getLabelBlockSpec(block, ctx);
        const labelHeight = Math.max(1, getLabelLines(block).length) * ctx.metrics.lineHeight;
        const lines = wrapRichParagraph(
          ctx.fonts,
          ctx.metrics.bodySize,
          block.textSegments || block.text,
          labelSpec.bodyWidth,
          block.firstLineIndent != null ? block.firstLineIndent : 0
        );
        const paragraphHeight = measureParagraphHeight(lines, ctx.metrics.lineHeight, ctx.spec.font.paragraphGap);
        height += Math.max(labelHeight, paragraphHeight - ctx.spec.font.paragraphGap) + labelSpec.gapAfter;
      } else if (block.type === 'labelLines') {
        const labelSpec = getLabelBlockSpec(block, ctx);
        const labelHeight = Math.max(1, getLabelLines(block).length) * ctx.metrics.lineHeight;
        let bodyHeight = 0;
        (block.lines || []).forEach((line) => {
          const lineEntry = normalizeLineEntry(line);
          const wrapped = wrapText(ctx.fonts.regular, ctx.metrics.bodySize, lineEntry.text, labelSpec.bodyWidth);
          bodyHeight += wrapped.length * ctx.metrics.lineHeight;
        });
        height += Math.max(labelHeight, bodyHeight || ctx.metrics.lineHeight) + labelSpec.gapAfter;
      } else if (block.type === 'lines') {
        const width = block.width != null
          ? block.width
          : ((ctx.spec.labelBlock && ctx.spec.labelBlock.bodyWidth) || ctx.spec.bodyStart.w);
        (block.lines || []).forEach((line) => {
          const lineEntry = normalizeLineEntry(line);
          const wrapped = wrapText(ctx.fonts.regular, ctx.metrics.bodySize, lineEntry.text, width);
          height += wrapped.length * ctx.metrics.lineHeight;
        });
        height += ctx.spec.font.paragraphGap;
      } else if (block.type === 'list') {
        height += ctx.spec.list.gapBefore || 0;
        (block.items || []).forEach((item, itemIndex) => {
          const itemLines = wrapRichParagraph(
            ctx.fonts,
            ctx.metrics.bodySize,
            item.textSegments || item.text,
            ctx.spec.list.contentWidth - ctx.spec.list.numberColumnWidth - 6
            ,
            0
          );
          height += itemLines.length * ctx.metrics.lineHeight;
          (item.lines || []).forEach((line) => {
            const lineEntry = normalizeLineEntry(line);
            const wrapped = wrapText(
              ctx.fonts.regular,
              ctx.metrics.bodySize,
              lineEntry.text,
              ctx.spec.list.contentWidth - ctx.spec.list.sublineIndent
            );
            height += wrapped.length * ctx.metrics.lineHeight;
          });
          if (itemIndex < (block.items || []).length - 1) {
            height += ctx.spec.list.itemGap;
          }
        });
        height += (ctx.spec.list.gapAfter != null ? ctx.spec.list.gapAfter : ctx.spec.font.paragraphGap);
      }
    });
    return height;
  }

  function drawParagraphBlock(page, ctx, block, topY) {
    const spec = ctx.spec;
    const width = block.width != null ? block.width : ((spec.labelBlock && spec.labelBlock.bodyWidth) || spec.bodyStart.w);
    const firstLineIndent = block.firstLineIndent != null ? block.firstLineIndent : spec.font.paragraphIndent;
    const baseX = block.x != null ? block.x : ((spec.labelBlock && spec.labelBlock.bodyX) || spec.bodyStart.x);
    const lines = wrapRichParagraph(ctx.fonts, ctx.metrics.bodySize, block.textSegments || block.text, width, firstLineIndent);
    let cursorY = topY;
    lines.forEach((line) => {
      drawRichSegmentsLine(page, ctx, line.segments || [{ text: line.text || '', bold:false }], baseX + (line.indent || 0), cursorY);
      cursorY += ctx.metrics.lineHeight;
    });
    return cursorY + spec.font.paragraphGap;
  }

  function drawLabelParagraphBlock(page, ctx, block, topY) {
    const labelSpec = getLabelBlockSpec(block, ctx);
    const labelLines = getLabelLines(block);
    const firstLineIndent = block.firstLineIndent != null ? block.firstLineIndent : 0;
    const lines = wrapRichParagraph(
      ctx.fonts,
      ctx.metrics.bodySize,
      block.textSegments || block.text,
      labelSpec.bodyWidth,
      firstLineIndent
    );
    const labelHeight = Math.max(1, labelLines.length) * ctx.metrics.lineHeight;
    const paragraphHeight = measureParagraphHeight(lines, ctx.metrics.lineHeight, ctx.spec.font.paragraphGap) - ctx.spec.font.paragraphGap;
    drawLabelStack(page, ctx, labelLines, labelSpec.labelX, topY, labelSpec.labelWidth);

    let cursorY = topY;
    lines.forEach((line) => {
      drawRichSegmentsLine(page, ctx, line.segments || [{ text: line.text || '', bold:false }], labelSpec.bodyX + (line.indent || 0), cursorY);
      cursorY += ctx.metrics.lineHeight;
    });
    return topY + Math.max(labelHeight, paragraphHeight) + labelSpec.gapAfter;
  }

  function drawLabelLinesBlock(page, ctx, block, topY) {
    const labelSpec = getLabelBlockSpec(block, ctx);
    const labelLines = getLabelLines(block);
    const labelHeight = Math.max(1, labelLines.length) * ctx.metrics.lineHeight;
    drawLabelStack(page, ctx, labelLines, labelSpec.labelX, topY, labelSpec.labelWidth);

    let cursorY = topY;
    let bodyHeight = 0;
    (block.lines || []).forEach((line) => {
      const lineEntry = normalizeLineEntry(line);
      const wrapped = wrapText(ctx.fonts.regular, ctx.metrics.bodySize, lineEntry.text, labelSpec.bodyWidth);
      wrapped.forEach((wrappedLine) => {
        page.drawText(wrappedLine, {
          x: labelSpec.bodyX,
          y: topToPdfY(page.getHeight(), cursorY, ctx.metrics.bodySize),
          size: ctx.metrics.bodySize,
          font: lineEntry.bold ? ctx.fonts.bold : ctx.fonts.regular,
          color: ctx.colors.black,
        });
        cursorY += ctx.metrics.lineHeight;
        bodyHeight += ctx.metrics.lineHeight;
      });
    });
    return topY + Math.max(labelHeight, bodyHeight || ctx.metrics.lineHeight) + labelSpec.gapAfter;
  }

  function drawLinesBlock(page, ctx, block, topY) {
    const spec = ctx.spec;
    let cursorY = topY;
    const x = block.x != null ? block.x : ((spec.labelBlock && spec.labelBlock.bodyX) || spec.bodyStart.x);
    const width = block.width != null ? block.width : ((spec.labelBlock && spec.labelBlock.bodyWidth) || spec.bodyStart.w);
    (block.lines || []).forEach((line) => {
      const lineEntry = normalizeLineEntry(line);
      const wrapped = wrapText(ctx.fonts.regular, ctx.metrics.bodySize, lineEntry.text, width);
      wrapped.forEach((wrappedLine) => {
        page.drawText(wrappedLine, {
          x,
          y: topToPdfY(page.getHeight(), cursorY, ctx.metrics.bodySize),
          size: ctx.metrics.bodySize,
          font: lineEntry.bold ? ctx.fonts.bold : ctx.fonts.regular,
          color: ctx.colors.black,
        });
        cursorY += ctx.metrics.lineHeight;
      });
    });
    return cursorY + spec.font.paragraphGap;
  }

  function drawListBlock(page, ctx, block, topY) {
    const spec = ctx.spec;
    let cursorY = topY + (spec.list.gapBefore || 0);
    (block.items || []).forEach((item, index) => {
      const itemNumber = `${index + 1}.`;
      page.drawText(itemNumber, {
        x: spec.list.startX,
        y: topToPdfY(page.getHeight(), cursorY, ctx.metrics.bodySize),
        size: ctx.metrics.bodySize,
        font: ctx.fonts.regular,
        color: ctx.colors.black,
      });

      const itemLines = wrapRichParagraph(
        ctx.fonts,
        ctx.metrics.bodySize,
        item.textSegments || item.text,
        spec.list.contentWidth - spec.list.numberColumnWidth - 6,
        0
      );

      itemLines.forEach((wrappedLine) => {
        drawRichSegmentsLine(
          page,
          ctx,
          wrappedLine.segments || [{ text: wrappedLine.text || '', bold:false }],
          spec.list.startX + spec.list.numberColumnWidth + 6 + (wrappedLine.indent || 0),
          cursorY
        );
        cursorY += ctx.metrics.lineHeight;
      });

      (item.lines || []).forEach((line) => {
        const lineEntry = normalizeLineEntry(line);
        const wrapped = wrapText(
          ctx.fonts.regular,
          ctx.metrics.bodySize,
          lineEntry.text,
          spec.list.contentWidth - spec.list.sublineIndent
        );
        wrapped.forEach((wrappedLine) => {
          page.drawText(wrappedLine, {
            x: spec.list.startX + spec.list.sublineIndent,
            y: topToPdfY(page.getHeight(), cursorY, ctx.metrics.bodySize),
            size: ctx.metrics.bodySize,
            font: lineEntry.bold ? ctx.fonts.bold : ctx.fonts.regular,
            color: ctx.colors.black,
          });
          cursorY += ctx.metrics.lineHeight;
        });
      });

      if (index < (block.items || []).length - 1) {
        cursorY += spec.list.itemGap;
      }
    });
    return cursorY + (spec.list.gapAfter != null ? spec.list.gapAfter : spec.font.paragraphGap);
  }

  function measureTail(payload, ctx) {
    const spec = ctx.spec;
    const closingParagraphs = payload.closingParagraphs || [];
    let closingHeight = 0;
    closingParagraphs.forEach((paragraph) => {
      const lines = wrapParagraphBalanced(
        ctx.fonts.regular,
        ctx.metrics.bodySize,
        paragraph,
        spec.closingBlock.w,
        spec.font.paragraphIndent
      );
      closingHeight += measureParagraphHeight(lines, ctx.metrics.lineHeight, spec.font.paragraphGap);
    });

    const signature = payload.signature || {};
    const showName = signature.showName === true;
    const signatureBlockDims = ctx.signatureBlockImage
      ? ctx.signatureBlockImage.scaleToFit(
          spec.signatureBlock.assetFitWidth,
          spec.signatureBlock.assetFitHeight
        )
      : null;
    const signatureHeight = signatureBlockDims
      ? signatureBlockDims.height
      : spec.signatureBlock.row1Height +
        (showName ? spec.signatureBlock.nameGap + 18 : 0) +
        spec.signatureBlock.positionGap +
        18;
    const contactHeight = (2 * ctx.metrics.lineHeight) + 2;

    return {
      closingHeight,
      signatureHeight,
      signatureBlockDims,
      contactHeight,
      totalHeight:
        spec.closingBlock.gapBefore +
        closingHeight +
        spec.signatureBlock.gapBefore +
        signatureHeight +
        spec.contactBlock.gapBefore +
        contactHeight,
    };
  }

  function drawTail(page, payload, ctx, currentY) {
    const spec = ctx.spec;
    const metrics = measureTail(payload, ctx);
    const closingTop = currentY + spec.closingBlock.gapBefore;
    const signatureTop = closingTop + metrics.closingHeight + spec.signatureBlock.gapBefore;
    const contactTop = spec.flowBottomLimit - metrics.contactHeight;

    if (signatureTop + metrics.signatureHeight + spec.contactBlock.gapBefore > contactTop) {
      throw new Error('ข้อมูลเกิน 1 หน้า กรุณาแยกหมาย');
    }

    let cursorY = closingTop;
    (payload.closingParagraphs || []).forEach((paragraph) => {
      const lines = wrapParagraphBalanced(
        ctx.fonts.regular,
        ctx.metrics.bodySize,
        paragraph,
        spec.closingBlock.w,
        spec.font.paragraphIndent
      );
      lines.forEach((line) => {
        if (!line.isBlank) {
          page.drawText(line.text || '', {
            x: spec.closingBlock.x + (line.indent || 0),
            y: topToPdfY(page.getHeight(), cursorY, ctx.metrics.bodySize),
            size: ctx.metrics.bodySize,
            font: ctx.fonts.regular,
            color: ctx.colors.black,
          });
        }
        cursorY += ctx.metrics.lineHeight;
      });
      cursorY += spec.font.paragraphGap;
    });

    const signature = payload.signature || {};
    const signatureCenterX = spec.signatureBlock.centerX || (spec.signatureBlock.x + (spec.signatureBlock.w / 2));

    if (ctx.signatureBlockImage && metrics.signatureBlockDims) {
      const dims = metrics.signatureBlockDims;
      page.drawImage(ctx.signatureBlockImage, {
        x: signatureCenterX - (dims.width / 2),
        y: topToPdfY(page.getHeight(), signatureTop, dims.height),
        width: dims.width,
        height: dims.height,
      });
    } else {
      const signatureText = normalizeSpace(signature.rank);
      const imageDims = ctx.signatureImage
        ? ctx.signatureImage.scaleToFit(
            spec.signatureBlock.imageFitWidth,
            spec.signatureBlock.imageFitHeight
          )
        : null;
      const textWidth = signatureText
        ? measureText(ctx.fonts.regular, ctx.metrics.bodySize, signatureText)
        : 0;
      const imageLeft = imageDims
        ? signatureCenterX - (imageDims.width / 2)
        : signatureCenterX;
      let line2MinX = 0;

      if (signatureText) {
        const rankX = imageDims
          ? imageLeft - spec.signatureBlock.groupGap - textWidth
          : signatureCenterX - (textWidth / 2);
        line2MinX = rankX + textWidth;
        page.drawText(signatureText, {
          x: rankX,
          y: topToPdfY(page.getHeight(), signatureTop + spec.signatureBlock.rankTopOffset, ctx.metrics.bodySize),
          size: ctx.metrics.bodySize,
          font: ctx.fonts.regular,
          color: ctx.colors.black,
        });
      }

      if (imageDims) {
        page.drawImage(ctx.signatureImage, {
          x: imageLeft,
          y: topToPdfY(page.getHeight(), signatureTop + spec.signatureBlock.imageTopOffset, imageDims.height),
          width: imageDims.width,
          height: imageDims.height,
        });
      }

      const showName = signature.showName === true;
      const nameTop = signatureTop + spec.signatureBlock.row1Height + spec.signatureBlock.nameGap;
      let positionTop = signatureTop + spec.signatureBlock.row1Height + spec.signatureBlock.positionGap;
      let positionCenterX = signatureCenterX;

      if (showName) {
        const nameText = `(${normalizeSpace(signature.name)})`;
        const nameWidth = measureText(ctx.fonts.regular, ctx.metrics.bodySize, nameText);
        const centeredNameX = signatureCenterX - (nameWidth / 2);
        const safeNameX = Math.max(centeredNameX, line2MinX);
        positionCenterX = safeNameX + (nameWidth / 2);
        page.drawText(nameText, {
          x: safeNameX,
          y: topToPdfY(page.getHeight(), nameTop, ctx.metrics.bodySize),
          size: ctx.metrics.bodySize,
          font: ctx.fonts.regular,
          color: ctx.colors.black,
        });
        positionTop = nameTop + 18 + spec.signatureBlock.positionGap;
      }

      const positionText = normalizeSpace(signature.position);
      const positionWidth = measureText(ctx.fonts.regular, ctx.metrics.bodySize, positionText);
      page.drawText(positionText, {
        x: positionCenterX - (positionWidth / 2),
        y: topToPdfY(page.getHeight(), positionTop, ctx.metrics.bodySize),
        size: ctx.metrics.bodySize,
        font: ctx.fonts.regular,
        color: ctx.colors.black,
      });
    }

    const contact = payload.contact || {};
    const contactLines = [normalizeSpace(contact.nameLine), normalizeSpace(contact.detailLine)].filter(Boolean);
    contactLines.forEach((line, index) => {
      page.drawText(line, {
        x: spec.contactBlock.x,
        y: topToPdfY(page.getHeight(), contactTop + (index * ctx.metrics.lineHeight), ctx.metrics.bodySize),
        size: ctx.metrics.bodySize,
        font: ctx.fonts.regular,
        color: ctx.colors.black,
      });
    });
  }

  function drawSharedFields(page, payload, ctx) {
    const spec = ctx.spec;
    const warnings = [];
    Object.keys(spec.sharedFields).forEach((key) => {
      const fieldSpec = spec.sharedFields[key];
      const value = normalizeSpace(payload.sharedFields && payload.sharedFields[key]);
      if (!value) return;

      const wrapped = wrapText(ctx.fonts.regular, fieldSpec.size, value, fieldSpec.w);
      const clamped = clampLines(ctx.fonts.regular, fieldSpec.size, wrapped, fieldSpec.w, fieldSpec.maxLines);
      if (clamped.truncated) {
        warnings.push(`ข้อความช่อง ${key} ยาวเกินกรอบ ระบบตัดข้อความท้ายออก`);
      }

      if (fieldSpec.clearBackground) {
        page.drawRectangle({
          x: fieldSpec.x - 2,
          y: topToPdfY(page.getHeight(), fieldSpec.y - 1, fieldSpec.h + 2),
          width: fieldSpec.w + 4,
          height: fieldSpec.h + 4,
          color: ctx.colors.white,
        });
      }

      drawSimpleLines(
        page,
        ctx.fonts.regular,
        fieldSpec.size,
        ctx.colors.black,
        clamped.lines,
        fieldSpec.x,
        fieldSpec.y,
        fieldSpec.w,
        spec.font.lineHeight,
        fieldSpec.align,
        fieldSpec.centerX
      );
    });
    return warnings;
  }

  async function renderDocument(payload) {
    assertDependencies();
    const spec = await resolveSpec(payload.config);

    const templateBytes = await fetchArrayBuffer(payload.templateUrl);
    const regularFontBytes = await fetchArrayBuffer(payload.fonts.regular);
    const boldFontBytes = await fetchArrayBuffer(payload.fonts.bold);
    const signatureBytes = payload.assets && payload.assets.signature
      ? await fetchArrayBuffer(payload.assets.signature)
      : null;
    const signatureBlockBytes = payload.assets && payload.assets.signatureBlock
      ? await fetchArrayBuffer(payload.assets.signatureBlock)
      : null;

    const pdfDoc = await PDFLibNS.PDFDocument.load(templateBytes);
    pdfDoc.registerFontkit(fontkitNS);

    const regularFont = await pdfDoc.embedFont(regularFontBytes);
    const boldFont = await pdfDoc.embedFont(boldFontBytes);
    const signatureImage = signatureBytes ? await pdfDoc.embedPng(signatureBytes) : null;
    const signatureBlockImage = signatureBlockBytes ? await pdfDoc.embedPng(signatureBlockBytes) : null;

    const page = pdfDoc.getPages()[0];
    const ctx = {
      fonts: {
        regular: regularFont,
        bold: boldFont,
      },
      colors: {
        black: PDFLibNS.rgb(COLOR_BLACK.r, COLOR_BLACK.g, COLOR_BLACK.b),
        white: PDFLibNS.rgb(COLOR_WHITE.r, COLOR_WHITE.g, COLOR_WHITE.b),
      },
      metrics: {
        bodySize: spec.font.size,
        lineHeight: spec.font.lineHeight,
      },
      signatureImage,
      signatureBlockImage,
      spec,
    };

    const warnings = drawSharedFields(page, payload, ctx);

    let currentY = spec.bodyStart.y;
    (payload.blocks || []).forEach((block) => {
      if (block.type === 'paragraph') {
        currentY = drawParagraphBlock(page, ctx, block, currentY);
        return;
      }
      if (block.type === 'labelParagraph') {
        currentY = drawLabelParagraphBlock(page, ctx, block, currentY);
        return;
      }
      if (block.type === 'labelLines') {
        currentY = drawLabelLinesBlock(page, ctx, block, currentY);
        return;
      }
      if (block.type === 'lines') {
        currentY = drawLinesBlock(page, ctx, block, currentY);
        return;
      }
      if (block.type === 'list') {
        currentY = drawListBlock(page, ctx, block, currentY);
      }
    });

    drawTail(page, payload, ctx, currentY);

    const bytes = await pdfDoc.save();
    return {
      bytes,
      blob: new Blob([bytes], { type: 'application/pdf' }),
      warnings,
    };
  }

  global.CSD1PdfTemplateRenderer = {
    renderDocument,
    defaultSpec: DEFAULT_SPEC,
  };
})(window);
