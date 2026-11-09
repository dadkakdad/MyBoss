/**
 * Download the deck as PNG or as an editable PowerPoint file.
 * Points on the 960×540 slide become inches at 72 pt per inch.
 */

import { cardStyle, displayName, headcountLabel, moderated, shapePerson } from "./org.js";
import { CARDS, LINE, LINK_STYLE, SLIDE, dottedDash } from "./layout.js";
import { contentFrame } from "./editor.js";

const STYLE = {
  ordinary: { fill: "001B58", color: "FFFFFF", bar: "00123F" },
  important: { fill: "8C0000", color: "FFFFFF", bar: "5C0000" },
  intern: { fill: "F4F4F4", color: "1A1A1A", bar: "FFFFFF" },
  split: { fill: "F4F4F4", color: "1A1A1A", bar: "F4F4F4", accent: "8C0000" },
  tbh: { fill: "FFFFFF", color: "001B57", bar: "F7F9FC" },
};

/** Type sizes for a card, the same as the on-screen card: the size's own pair, scaled with the card's height. */
function cardType(shape) {
  const spec = CARDS[shape.size] || CARDS.Medium;
  const k = spec.h ? shape.h / spec.h : 1;
  return { name: spec.name * k, job: spec.title * k };
}

function personLines(org, shape) {
  if (!shape.key) return shape.lines || ["", "", ""];
  const person = shapePerson(org, shape);
  if (!person) return ["", "", ""];
  if (person.tbh) return shape.h >= 32 ? ["To be hired", "", moderated(person.job)] : ["To be hired", moderated(person.job), ""];
  return [moderated(person.surname), moderated(person.name), moderated(person.job)];
}

export function slideToCanvas(org, slide, scale = 2) {
  const frame = contentFrame(slide.shapes || []);
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.ceil(frame.w * scale));
  canvas.height = Math.max(1, Math.ceil(frame.h * scale));
  const ctx = canvas.getContext("2d");
  ctx.scale(scale, scale);
  ctx.translate(-frame.x, -frame.y);
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(frame.x, frame.y, frame.w, frame.h);
  const shapes = [...slide.shapes].sort((a, b) => (a.z || 0) - (b.z || 0));
  for (const shape of shapes) drawShape(ctx, org, shape);
  return canvas;
}

/** The text cut to `maxW` with an ellipsis, using the font currently set on ctx. */
function fitText(ctx, text, maxW) {
  const value = String(text ?? "");
  if (!value || ctx.measureText(value).width <= maxW) return value;
  let cut = value;
  while (cut.length > 1 && ctx.measureText(`${cut}…`).width > maxW) cut = cut.slice(0, -1);
  return `${cut.trimEnd()}…`;
}

/** Word-wrapped lines no wider than `maxW`; the last allowed line ends in an ellipsis when text is left over. */
function wrapLines(ctx, text, maxW, maxLines) {
  const words = String(text ?? "").split(/\s+/).filter(Boolean);
  const lines = [];
  let line = "";
  for (const word of words) {
    const next = line ? `${line} ${word}` : word;
    if (line && ctx.measureText(next).width > maxW) {
      lines.push(line);
      line = word;
    } else line = next;
  }
  if (line) lines.push(line);
  if (lines.length > maxLines) {
    const kept = lines.slice(0, maxLines);
    kept[maxLines - 1] = fitText(ctx, `${kept[maxLines - 1]} ${lines.slice(maxLines).join(" ")}`, maxW);
    return kept;
  }
  return lines.map((item) => fitText(ctx, item, maxW));
}

function drawShape(ctx, org, shape) {
  if (shape.type === "shade") {
    ctx.fillStyle = shape.fill || "#DEEBF7";
    ctx.fillRect(shape.x, shape.y, shape.w, shape.h);
    ctx.fillStyle = "#173C5D";
    ctx.font = "11px Arial";
    ctx.textAlign = "center";
    ctx.textBaseline = "bottom";
    ctx.fillText(fitText(ctx, shape.text || "", shape.w - 8), shape.x + shape.w / 2, shape.y + shape.h - 5);
    ctx.textAlign = "left";
    ctx.textBaseline = "alphabetic";
    return;
  }
  if (shape.type === "wire") {
    ctx.beginPath();
    ctx.moveTo(shape.x1, shape.y1);
    ctx.lineTo(shape.x2, shape.y2);
    ctx.strokeStyle = shape.dotted ? LINE.dotted : LINE.solid;
    ctx.lineWidth = shape.weight || 2;
    ctx.lineCap = shape.dotted ? "round" : "square";
    ctx.setLineDash(shape.dotted ? dottedDash(shape.weight || 2) : []);
    ctx.stroke();
    ctx.setLineDash([]);
    return;
  }
  if (shape.type === "frame") {
    ctx.strokeStyle = "#9d1c2a";
    ctx.lineWidth = 1.5;
    ctx.setLineDash([5, 3]);
    ctx.strokeRect(shape.x, shape.y, shape.w, shape.h);
    ctx.setLineDash([]);
    return;
  }
  if (shape.type === "title") {
    ctx.fillStyle = "#001B58";
    ctx.font = "bold 26px Arial";
    ctx.fillText(shape.name || "", shape.x, shape.y + 26);
    ctx.font = "16px Arial";
    ctx.fillStyle = "#3d4d66";
    ctx.fillText(shape.department || "Organisation", shape.x, shape.y + 48);
    return;
  }
  if (shape.type === "stats") {
    ctx.fillStyle = "#f3f3f3";
    ctx.strokeStyle = "#b9b9b9";
    ctx.lineWidth = 1;
    roundRect(ctx, shape.x, shape.y, shape.w, shape.h, 4);
    ctx.fill();
    ctx.stroke();
    ctx.fillStyle = "#1c2430";
    ctx.font = "bold 11px Arial";
    ctx.fillText(shape.heading || "Department statistics", shape.x + 8, shape.y + 14);
    ctx.font = "10px Arial";
    (shape.lines || []).forEach((line, index) => {
      ctx.fillText(`• ${line}`, shape.x + 8, shape.y + 30 + index * 12);
    });
    return;
  }
  if (shape.type === "hiddenlist") {
    ctx.fillStyle = "#8A8F98";
    ctx.font = "8px Arial";
    const words = String(shape.text || "").split(" ");
    let line = "";
    let row = 0;
    for (const word of words) {
      const next = line ? `${line} ${word}` : word;
      if (line && ctx.measureText(next).width > shape.w) {
        ctx.fillText(line, shape.x, shape.y + 8 + row * 10);
        row += 1;
        line = word;
      } else line = next;
    }
    if (line) ctx.fillText(line, shape.x, shape.y + 8 + row * 10);
    return;
  }
  if (shape.type === "legend") {
    ctx.fillStyle = "#5c6778";
    ctx.font = "bold 10px Arial";
    ctx.fillText("Indicators", shape.x, shape.y + 10);
    const rows = [
      ["ordinary", "Ordinary employee"],
      ["important", "Important employee"],
      ["intern", "Intern"],
    ];
    rows.forEach(([kind, label], index) => {
      ctx.fillStyle = `#${STYLE[kind].fill}`;
      ctx.fillRect(shape.x, shape.y + 16 + index * 14, 12, 8);
      ctx.fillStyle = "#1c2430";
      ctx.font = "10px Arial";
      ctx.fillText(label, shape.x + 16, shape.y + 24 + index * 14);
    });
    return;
  }
  if (shape.type === "link") {
    ctx.fillStyle = LINK_STYLE.fill;
    ctx.fillRect(shape.x, shape.y, shape.w, shape.h);
    ctx.fillStyle = LINK_STYLE.color;
    ctx.font = "bold 8px Arial";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText(fitText(ctx, shape.text || "", shape.w - 4), shape.x + shape.w / 2, shape.y + shape.h / 2 + 0.3);
    ctx.textAlign = "left";
    ctx.textBaseline = "alphabetic";
    return;
  }
  if (shape.type === "nav") {
    ctx.fillStyle = "#173C5D";
    ctx.font = "11px Arial";
    ctx.textBaseline = "middle";
    const text = fitText(ctx, shape.text || "", shape.w);
    ctx.fillText(text, shape.x, shape.y + shape.h / 2);
    const width = ctx.measureText(text).width;
    ctx.fillRect(shape.x, shape.y + shape.h / 2 + 6, width, 0.8);
    ctx.textBaseline = "alphabetic";
    return;
  }
  if (shape.type === "bubble" || shape.type === "comment") {
    ctx.fillStyle = "#ffffff";
    ctx.strokeStyle = shape.dotted ? "#00123F" : "#8a8175";
    ctx.lineWidth = 1;
    ctx.setLineDash(shape.dotted ? [3, 2] : []);
    roundRect(ctx, shape.x, shape.y, shape.w, shape.h, 3);
    ctx.fill();
    ctx.stroke();
    ctx.setLineDash([]);
    ctx.fillStyle = "#001B58";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    const cx = shape.x + shape.w / 2;
    if (shape.type === "comment") {
      // Same look as on screen: bold 7 pt, wrapped inside the box, centred.
      ctx.font = "bold 7px Arial";
      const lines = wrapLines(ctx, shape.text, shape.w - 6, Math.max(1, Math.floor(shape.h / 8)));
      const step = 8;
      const top = shape.y + shape.h / 2 - ((lines.length - 1) * step) / 2;
      lines.forEach((line, index) => ctx.fillText(line, cx, top + index * step + 0.3));
    } else {
      ctx.font = "bold 8px Arial";
      ctx.fillText(fitText(ctx, shape.text || "", shape.w - 6), cx, shape.y + shape.h / 2 + 0.3);
    }
    ctx.textAlign = "left";
    ctx.textBaseline = "alphabetic";
    return;
  }
  if (shape.type === "cc") {
    ctx.fillStyle = "#00ff00";
    ctx.strokeStyle = "#0b6b0b";
    ctx.lineWidth = 1;
    ctx.fillRect(shape.x, shape.y, shape.w, shape.h);
    ctx.strokeRect(shape.x, shape.y, shape.w, shape.h);
    ctx.fillStyle = "#083008";
    ctx.font = `bold ${shape.fs || Math.max(5, shape.h * 0.64)}px Arial`;
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText(shape.text || "?", shape.x + shape.w / 2, shape.y + shape.h / 2);
    ctx.textAlign = "left";
    ctx.textBaseline = "alphabetic";
    return;
  }
  if (shape.type === "card") {
    const kind = shape.key ? cardStyle(shapePerson(org, shape)) : shape.placeholder || "ordinary";
    const palette = STYLE[kind];
    const vacant = shape.key ? !!shapePerson(org, shape)?.tbh : kind === "tbh";
    ctx.save();
    roundRect(ctx, shape.x, shape.y, shape.w, shape.h, 2);
    ctx.clip();
    ctx.fillStyle = `#${palette.fill}`;
    ctx.fillRect(shape.x, shape.y, shape.w, shape.h);
    if (kind === "split") {
      ctx.fillStyle = `#${palette.accent}`;
      ctx.beginPath();
      ctx.moveTo(shape.x, shape.y);
      ctx.lineTo(shape.x + shape.w, shape.y);
      ctx.lineTo(shape.x + shape.w, shape.y + shape.h);
      ctx.closePath();
      ctx.fill();
      ctx.beginPath();
      ctx.moveTo(shape.x, shape.y);
      ctx.lineTo(shape.x + shape.w, shape.y + shape.h);
      ctx.strokeStyle = "#1a1a1a";
      ctx.lineWidth = 1.25;
      ctx.stroke();
    }
    ctx.restore();
    if (kind === "split") {
      ctx.strokeStyle = "#1a1a1a";
      ctx.lineWidth = 1;
      ctx.strokeRect(shape.x, shape.y, shape.w, shape.h);
    }
    if (vacant) {
      ctx.save();
      ctx.strokeStyle = kind === "ordinary" || kind === "important" ? "#ffffff" : "#001B57";
      ctx.lineWidth = 1;
      ctx.setLineDash([3, 2]);
      ctx.strokeRect(shape.x + 2.5, shape.y + 2.5, shape.w - 5, shape.h - 5);
      ctx.restore();
    }
    const lines = personLines(org, shape);
    const person = shapePerson(org, shape);
    const footer = person?.location || "";
    const { name: nameSize, job: jobSize } = cardType(shape);
    const padX = Math.max(4, Math.round(nameSize * 0.65));
    const padY = Math.round(padX * 0.75);
    const textW = shape.w - padX * 2;
    // Every line is cut to the card with an ellipsis and clipped to it, as the on-screen card does.
    ctx.save();
    ctx.beginPath();
    ctx.rect(shape.x, shape.y, shape.w, shape.h);
    ctx.clip();
    ctx.textBaseline = "top";
    ctx.fillStyle = `#${palette.color}`;
    ctx.font = `bold ${nameSize}px Arial`;
    let top = shape.y + padY;
    for (const line of lines.slice(0, 2)) {
      if (!line) continue; // an empty line takes no room, so a vacancy's job sits right under "To be hired"
      ctx.fillText(fitText(ctx, line, textW), shape.x + padX, top);
      top += nameSize * 1.05;
    }
    if (shape.h >= 32 && lines[2]) {
      ctx.font = `${jobSize}px Arial`;
      ctx.fillText(fitText(ctx, lines[2], textW), shape.x + padX, top + 1);
    }
    const badge = person ? headcountLabel(person) : "";
    ctx.textBaseline = "middle";
    if (footer) {
      ctx.fillStyle = `#${palette.bar}`;
      ctx.fillRect(shape.x, shape.y + shape.h - 11, shape.w, 11);
      ctx.fillStyle = kind === "ordinary" || kind === "important" ? "#ffffff" : "#1c2430";
      let room = textW;
      if (badge) {
        ctx.font = "bold 7px Arial";
        ctx.textAlign = "right";
        ctx.fillText(badge, shape.x + shape.w - padX, shape.y + shape.h - 5.5);
        room -= ctx.measureText(badge).width + 4;
        ctx.textAlign = "left";
      }
      ctx.font = "7px Arial";
      ctx.fillText(fitText(ctx, footer, room), shape.x + padX, shape.y + shape.h - 5.5);
    } else if (badge) {
      ctx.font = "bold 8px Arial";
      ctx.fillStyle = kind === "split" ? "#ffffff" : `#${palette.color}`;
      ctx.textAlign = "right";
      ctx.fillText(badge, shape.x + shape.w - padX, shape.y + shape.h - 6);
      ctx.textAlign = "left";
    }
    ctx.restore();
    if (person?.newcomer) drawNewBadge(ctx, shape, person);
  }
}

/** Badge size follows the card's name type, like the on-screen badge (3.4 em of 0.8 of the name size; "New" and the date on two lines). */
function newBadgeBox(shape, person) {
  const spec = CARDS[shape.size] || CARDS.Medium;
  const u = 0.8 * spec.name * (spec.h ? shape.h / spec.h : 1);
  const lines = person?.newcomerDate ? 2 : 1;
  return { u, w: u * 3.4, h: u * (0.3 + 1.15 * lines), x: shape.x + shape.w - 3 - u * 3.4, y: shape.y + 2, dated: lines === 2 };
}

function drawNewBadge(ctx, shape, person) {
  const box = newBadgeBox(shape, person);
  ctx.save();
  roundRect(ctx, box.x, box.y, box.w, box.h, 2);
  ctx.fillStyle = "#FFC61A";
  ctx.fill();
  ctx.fillStyle = "#3D2A00";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.font = `bold ${box.u}px Arial`;
  const cx = box.x + box.w / 2;
  if (box.dated) {
    ctx.fillText("New", cx, box.y + box.u * 0.75);
    ctx.font = `600 ${box.u * 0.9}px Arial`;
    ctx.fillText(person.newcomerDate, cx, box.y + box.u * 1.9);
  } else ctx.fillText("New", cx, box.y + box.h / 2 + 0.3);
  ctx.restore();
}

function roundRect(ctx, x, y, w, h, r) {
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.arcTo(x + w, y, x + w, y + h, r);
  ctx.arcTo(x + w, y + h, x, y + h, r);
  ctx.arcTo(x, y + h, x, y, r);
  ctx.arcTo(x, y, x + w, y, r);
  ctx.closePath();
}

function inches(points) {
  return points / 72;
}

function addLine(slide, shape) {
  const x1 = inches(shape.x1);
  const y1 = inches(shape.y1);
  const x2 = inches(shape.x2);
  const y2 = inches(shape.y2);
  slide.addShape("line", {
    x: Math.min(x1, x2),
    y: Math.min(y1, y2),
    w: Math.max(Math.abs(x2 - x1), 0.01),
    h: Math.max(Math.abs(y2 - y1), 0.01),
    flipH: x1 > x2,
    flipV: y1 > y2,
    line: {
      color: (shape.dotted ? LINE.dotted : LINE.solid).replace("#", ""),
      width: shape.weight || 1.5,
      dashType: shape.dotted ? "sysDot" : "solid",
    },
  });
}

export async function deckToPptx(org, slides, fileName) {
  const PptxGenJS = window.PptxGenJS;
  if (!PptxGenJS) throw new Error("PowerPoint export is not loaded.");
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: "MYBOSS", width: inches(SLIDE.width), height: inches(SLIDE.height) });
  pptx.layout = "MYBOSS";
  pptx.author = "MyBoss";
  pptx.title = fileName || "Org chart";

  slides.forEach((source, index) => {
    const slide = pptx.addSlide();
    slide.background = { color: "FFFFFF" };
    const shapes = [...source.shapes].sort((a, b) => (a.z || 0) - (b.z || 0));
    for (const shape of shapes) {
      if (shape.type === "wire") {
        addLine(slide, shape);
        continue;
      }
      if (shape.type === "shade") {
        slide.addShape("rect", {
          x: inches(shape.x),
          y: inches(shape.y),
          w: inches(shape.w),
          h: inches(shape.h),
          fill: { color: (shape.fill || "#DEEBF7").replace("#", "") },
        });
        slide.addText(shape.text || "", {
          x: inches(shape.x + 4),
          y: inches(shape.y + shape.h - 16),
          w: inches(shape.w - 8),
          h: inches(14),
          fontFace: "Arial",
          fontSize: 11,
          color: "173C5D",
          align: "center",
        });
        continue;
      }
      if (shape.type === "title") {
        slide.addText(
          [
            { text: `${shape.name || ""}\n`, options: { fontSize: 26, bold: true, color: "001B58" } },
            { text: shape.department || "Organisation", options: { fontSize: 16, color: "3D4D66" } },
          ],
          { x: inches(shape.x), y: inches(shape.y), w: inches(shape.w), h: inches(shape.h), fontFace: "Arial", margin: 0 },
        );
        continue;
      }
      if (shape.type === "stats") {
        slide.addText(
          [{ text: `${shape.heading || "Department statistics"}\n`, options: { bold: true, fontSize: 11 } }, ...(shape.lines || []).map((line) => ({ text: `${line}\n`, options: { fontSize: 10 } }))],
          {
            x: inches(shape.x),
            y: inches(shape.y),
            w: inches(shape.w),
            h: inches(shape.h),
            fontFace: "Arial",
            color: "1C2430",
            fill: { color: "F3F3F3" },
            valign: "top",
          },
        );
        continue;
      }
      if (shape.type === "hiddenlist") {
        slide.addText(shape.text || "", {
          x: inches(shape.x),
          y: inches(shape.y),
          w: inches(shape.w),
          h: inches(shape.h),
          fontFace: "Arial",
          fontSize: 8,
          color: "8A8F98",
          margin: 0,
          valign: "top",
        });
        continue;
      }
      if (shape.type === "legend") {
        slide.addText("Indicators\nOrdinary employee\nImportant employee\nIntern", {
          x: inches(shape.x),
          y: inches(shape.y),
          w: inches(shape.w),
          h: inches(shape.h),
          fontFace: "Arial",
          fontSize: 10,
          color: "1C2430",
        });
        continue;
      }
      if (shape.type === "card") {
        const kind = shape.key ? cardStyle(shapePerson(org, shape)) : shape.placeholder || "ordinary";
        const palette = STYLE[kind];
        const vacant = shape.key ? !!shapePerson(org, shape)?.tbh : kind === "tbh";
        const lines = personLines(org, shape);
        const person = shapePerson(org, shape);
        const footer = person?.location || "";
        const { name: nameSize, job: jobSize } = cardType(shape);
        slide.addShape("rect", {
          x: inches(shape.x),
          y: inches(shape.y),
          w: inches(shape.w),
          h: inches(shape.h),
          fill: { color: palette.fill },
          line: vacant
            ? {
                color: kind === "ordinary" || kind === "important" ? "FFFFFF" : "001B57",
                width: 1,
                dashType: "dash",
              }
            : { color: palette.fill },
        });
        if (kind === "split") {
          slide.addShape("custGeom", {
            x: inches(shape.x),
            y: inches(shape.y),
            w: inches(shape.w),
            h: inches(shape.h),
            fill: { color: palette.accent },
            line: { color: palette.accent, width: 0 },
            points: [
              { x: 0, y: 0, moveTo: true },
              { x: inches(shape.w), y: 0 },
              { x: inches(shape.w), y: inches(shape.h) },
              { close: true },
            ],
          });
          slide.addShape("line", {
            x: inches(shape.x),
            y: inches(shape.y),
            w: inches(shape.w),
            h: inches(shape.h),
            line: { color: "1A1A1A", width: 1.25 },
          });
        }
        const badge = person ? headcountLabel(person) : "";
        const textRuns = [
          { text: lines[0] || "", options: { bold: true, fontSize: nameSize, breakLine: true } },
          { text: lines[1] || "", options: { bold: true, fontSize: nameSize, breakLine: true } },
        ];
        if (shape.h >= 32 && lines[2]) textRuns.push({ text: lines[2], options: { bold: false, fontSize: jobSize } });
        slide.addText(textRuns, {
          x: inches(shape.x + 5),
          y: inches(shape.y + 2),
          w: inches(shape.w - 10),
          h: inches(shape.h - (footer ? 12 : 2)),
          fontFace: "Arial",
          color: palette.color,
          margin: 0,
          valign: "top",
        });
        if (footer) {
          slide.addShape("rect", {
            x: inches(shape.x),
            y: inches(shape.y + shape.h - 11),
            w: inches(shape.w),
            h: inches(11),
            fill: { color: palette.bar },
          });
          slide.addText(footer, {
            x: inches(shape.x + 3),
            y: inches(shape.y + shape.h - 11),
            w: inches(Math.max(8, shape.w - (badge ? 32 : 6))),
            h: inches(11),
            fontFace: "Arial",
            fontSize: 7,
            color: kind === "ordinary" || kind === "important" ? "FFFFFF" : "1C2430",
            margin: 0,
          });
          if (badge) {
            slide.addText(badge, {
              x: inches(shape.x + shape.w - 30),
              y: inches(shape.y + shape.h - 11),
              w: inches(26),
              h: inches(11),
              align: "right",
              fontFace: "Arial",
              fontSize: 7,
              bold: true,
              color: kind === "ordinary" || kind === "important" ? "FFFFFF" : "1C2430",
              margin: 0,
            });
          }
        } else if (badge) {
          slide.addText(badge, {
            x: inches(shape.x),
            y: inches(shape.y + shape.h - 11),
            w: inches(shape.w - 4),
            h: inches(10),
            align: "right",
            valign: "bottom",
            fontFace: "Arial",
            fontSize: 8,
            bold: true,
            color: kind === "split" ? "FFFFFF" : palette.color,
          });
        }
        if (person?.newcomer) {
          const badgeBox = newBadgeBox(shape, person);
          slide.addText(
            person.newcomerDate
              ? [{ text: "New\n", options: { fontSize: badgeBox.u, bold: true } }, { text: person.newcomerDate, options: { fontSize: badgeBox.u * 0.9, bold: false } }]
              : "New",
            {
              x: inches(badgeBox.x),
              y: inches(badgeBox.y),
              w: inches(badgeBox.w),
              h: inches(badgeBox.h),
              fontFace: "Arial",
              fontSize: badgeBox.u,
              bold: true,
              align: "center",
              valign: "middle",
              color: "3D2A00",
              fill: { color: "FFC61A" },
              margin: 0,
            },
          );
        }
        continue;
      }
      if (shape.type === "cc") {
        slide.addText(shape.text || "?", {
          x: inches(shape.x),
          y: inches(shape.y),
          w: inches(shape.w),
          h: inches(shape.h),
          fontFace: "Arial",
          fontSize: shape.fs ? Math.round(shape.fs * 10) / 10 : Math.max(5, Math.round(shape.h * 0.64)),
          bold: true,
          align: "center",
          valign: "middle",
          color: "083008",
          fill: { color: "00FF00" },
          margin: 0,
        });
        continue;
      }
      if (shape.type === "link" || shape.type === "nav" || shape.type === "bubble" || shape.type === "comment") {
        const link = shape.type === "link";
        slide.addText(shape.text || "", {
          x: inches(shape.x),
          y: inches(shape.y),
          w: inches(Math.max(shape.w, 8)),
          h: inches(Math.max(shape.h, 8)),
          fontFace: "Arial",
          fontSize: shape.type === "comment" ? 7 : 8,
          bold: link,
          color: link ? LINK_STYLE.color.replace("#", "") : shape.type === "nav" ? "FFFFFF" : "001B58",
          fill: { color: link ? LINK_STYLE.fill.replace("#", "") : shape.type === "nav" ? "173C5D" : "FFFFFF" },
          align: "center",
          valign: "middle",
          wrap: shape.type === "comment",
          hyperlink: shape.target ? { slide: shape.target } : undefined,
        });
      }
    }
    const root = source.rootLabel || (source.rootKey && org.people.get(source.rootKey) ? displayName(org.people.get(source.rootKey)) : "");
    const caption = source.title || (root ? `${root} · slide ${index + 1}` : `Slide ${index + 1}`);
    slide.addNotes(caption);
  });

  await pptx.writeFile({ fileName: fileName || "MyBoss-org-chart.pptx" });
}
