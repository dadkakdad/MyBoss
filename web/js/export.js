/**
 * Download the deck as PNG or as an editable PowerPoint file.
 * Points on the 960×540 slide become inches at 72 pt per inch.
 */

import { cardStyle, displayName, headcountLabel, moderated } from "./org.js";
import { LINE, LINK_STYLE, SLIDE, dottedDash } from "./layout.js";
import { contentFrame } from "./editor.js";

const STYLE = {
  ordinary: { fill: "001B58", color: "FFFFFF", bar: "00123F" },
  important: { fill: "8C0000", color: "FFFFFF", bar: "5C0000" },
  outsource: { fill: "F4F4F4", color: "1A1A1A", bar: "FFFFFF" },
  split: { fill: "F4F4F4", color: "1A1A1A", bar: "F4F4F4", accent: "8C0000" },
  tbh: { fill: "FFFFFF", color: "001B57", bar: "F7F9FC" },
};

function personLines(org, shape) {
  if (!shape.key) return shape.lines || ["", "", ""];
  const person = org.people.get(shape.key);
  if (!person) return ["", "", ""];
  if (person.tbh) return ["To be hired", moderated(person.job), ""];
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

function drawShape(ctx, org, shape) {
  if (shape.type === "shade") {
    ctx.fillStyle = shape.fill || "#DEEBF7";
    ctx.fillRect(shape.x, shape.y, shape.w, shape.h);
    ctx.fillStyle = "#173C5D";
    ctx.font = "9px Arial";
    ctx.fillText(shape.text || "", shape.x + 6, shape.y + 12);
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
  if (shape.type === "legend") {
    ctx.fillStyle = "#5c6778";
    ctx.font = "bold 10px Arial";
    ctx.fillText("Indicators", shape.x, shape.y + 10);
    const rows = [
      ["ordinary", "Ordinary employee"],
      ["important", "Important employee"],
      ["outsource", "Outsourcer"],
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
    ctx.fillText(shape.text || "", shape.x + shape.w / 2, shape.y + shape.h / 2);
    ctx.textAlign = "left";
    ctx.textBaseline = "alphabetic";
    return;
  }
  if (shape.type === "bubble" || shape.type === "comment" || shape.type === "nav") {
    ctx.fillStyle = shape.type === "nav" ? "#173C5D" : "#ffffff";
    ctx.strokeStyle = shape.dotted ? "#00123F" : "#8a8175";
    ctx.lineWidth = 1;
    ctx.setLineDash(shape.dotted ? [3, 2] : []);
    roundRect(ctx, shape.x, shape.y, shape.w, shape.h, 3);
    ctx.fill();
    ctx.stroke();
    ctx.setLineDash([]);
    ctx.fillStyle = shape.type === "nav" ? "#ffffff" : "#1c2430";
    ctx.font = "8px Arial";
    ctx.fillText(shape.text || "", shape.x + 4, shape.y + shape.h - 3);
    return;
  }
  if (shape.type === "cc") {
    ctx.fillStyle = "#00ff00";
    ctx.strokeStyle = "#0b6b0b";
    ctx.lineWidth = 1;
    ctx.fillRect(shape.x, shape.y, shape.w, shape.h);
    ctx.strokeRect(shape.x, shape.y, shape.w, shape.h);
    ctx.fillStyle = "#083008";
    ctx.font = `bold ${Math.max(6, shape.h * (10 / 30))}px Arial`;
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText(shape.text || "?", shape.x + shape.w / 2, shape.y + shape.h / 2);
    ctx.textAlign = "left";
    ctx.textBaseline = "alphabetic";
    return;
  }
  if (shape.type === "card") {
    const kind = shape.key ? cardStyle(org.people.get(shape.key)) : shape.placeholder || "ordinary";
    const palette = STYLE[kind];
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
    ctx.strokeStyle = kind === "tbh" ? "#001B57" : "transparent";
    ctx.lineWidth = 1;
    if (kind === "tbh") {
      roundRect(ctx, shape.x, shape.y, shape.w, shape.h, 2);
      ctx.stroke();
    }
    const lines = personLines(org, shape);
    const person = shape.key ? org.people.get(shape.key) : null;
    const footer = person?.location || "";
    const nameSize = shape.h < 30 ? 7 : shape.h < 40 ? 8 : shape.h < 48 ? 10 : 12;
    ctx.fillStyle = `#${palette.color}`;
    ctx.font = `bold ${nameSize}px Arial`;
    ctx.fillText(lines[0] || "", shape.x + 4, shape.y + nameSize + 2);
    ctx.fillText(lines[1] || "", shape.x + 4, shape.y + nameSize * 2 + 3);
    if (shape.h >= 32) {
      ctx.font = `${Math.max(6, nameSize - 3)}px Arial`;
      ctx.fillText(lines[2] || "", shape.x + 4, shape.y + shape.h - (footer ? 12 : 4));
    }
    const badge = person ? headcountLabel(person) : "";
    if (footer) {
      ctx.fillStyle = `#${palette.bar}`;
      ctx.fillRect(shape.x, shape.y + shape.h - 11, shape.w, 11);
      ctx.fillStyle = kind === "ordinary" || kind === "important" ? "#ffffff" : "#1c2430";
      ctx.font = "7px Arial";
      ctx.fillText(footer, shape.x + 4, shape.y + shape.h - 3);
      if (badge) {
        ctx.font = "bold 7px Arial";
        ctx.textAlign = "right";
        ctx.fillText(badge, shape.x + shape.w - 4, shape.y + shape.h - 3);
        ctx.textAlign = "left";
      }
    } else if (badge) {
      ctx.font = "bold 8px Arial";
      ctx.fillStyle = kind === "split" ? "#ffffff" : `#${palette.color}`;
      ctx.textAlign = "right";
      ctx.fillText(badge, shape.x + shape.w - 4, shape.y + 9);
      ctx.textAlign = "left";
    }
  }
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
      if (shape.type === "legend") {
        slide.addText("Indicators\nOrdinary employee\nImportant employee\nOutsourcer", {
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
        const kind = shape.key ? cardStyle(org.people.get(shape.key)) : shape.placeholder || "ordinary";
        const palette = STYLE[kind];
        const lines = personLines(org, shape);
        const person = shape.key ? org.people.get(shape.key) : null;
        const footer = person?.location || "";
        const nameSize = shape.size && { ExtraSmall: 7, Small: 7, Medium: 9, Large: 11 }[shape.size];
        slide.addShape("rect", {
          x: inches(shape.x),
          y: inches(shape.y),
          w: inches(shape.w),
          h: inches(shape.h),
          fill: { color: palette.fill },
          line: kind === "tbh" ? { color: "001B57", width: 1 } : { color: palette.fill },
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
        slide.addText(
          [lines[0], lines[1], shape.h >= 32 ? lines[2] : ""].filter(Boolean).join("\n"),
          {
            x: inches(shape.x + 3),
            y: inches(shape.y + 1),
            w: inches(shape.w - 6),
            h: inches(shape.h - (footer ? 12 : 2)),
            fontFace: "Arial",
            fontSize: nameSize || (shape.h < 32 ? 7 : 9),
            color: palette.color,
            bold: true,
            margin: 0,
          },
        );
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
            y: inches(shape.y),
            w: inches(shape.w - 2),
            h: inches(10),
            align: "right",
            fontFace: "Arial",
            fontSize: 8,
            bold: true,
            color: kind === "split" ? "FFFFFF" : palette.color,
          });
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
          fontSize: 10,
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
    const root = source.rootKey ? org.people.get(source.rootKey) : null;
    const caption = source.title || (root ? `${displayName(root)} · slide ${index + 1}` : `Slide ${index + 1}`);
    slide.addNotes(caption);
  });

  await pptx.writeFile({ fileName: fileName || "MyBoss-org-chart.pptx" });
}
