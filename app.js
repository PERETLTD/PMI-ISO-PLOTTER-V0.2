const canvas = document.getElementById("plotCanvas");
const ctx = canvas.getContext("2d");
const canvasStage = document.getElementById("canvasStage");
const annotationEditor = document.getElementById("annotationEditor");

const strokeColorInput = document.getElementById("strokeColor");
const strokeWidthInput = document.getElementById("strokeWidth");
const strokeWidthValue = document.getElementById("strokeWidthValue");
const lineModeButton = document.getElementById("lineModeButton");
const textModeButton = document.getElementById("textModeButton");
const textControls = document.getElementById("textControls");
const textColorInput = document.getElementById("textColor");
const fontSizeInput = document.getElementById("fontSize");
const fontSizeValue = document.getElementById("fontSizeValue");
const fontFamilyInput = document.getElementById("fontFamily");
const showNodesInput = document.getElementById("showNodes");
const showGuidesInput = document.getElementById("showGuides");
const undoButton = document.getElementById("undoButton");
const clearButton = document.getElementById("clearButton");
const exportButton = document.getElementById("exportButton");

const page = {
  width: canvas.width,
  height: canvas.height,
  marginLeft: 82,
  marginTop: 72,
  gridWidth: 1031,
  gridHeight: 1474
};

const step = {
  vertical: 56.6929134,
  diagonalX: 49.098,
  diagonalY: 28.3464567
};

const state = {
  segments: [],
  annotations: [],
  pendingPoint: null,
  lineChainPoint: null,
  hoverPoint: null,
  tool: "line",
  history: [],
  activeAnnotationId: null,
  editingAnnotationId: null,
  rotatingAnnotationId: null,
  draggingAnnotationId: null,
  pendingDragAnnotationId: null,
  pointerDownPoint: null,
  dragOffset: { x: 0, y: 0 },
  nextAnnotationId: 1
};

const logoImage = new Image();
logoImage.src = "./pmi001-2.png";
logoImage.addEventListener("load", redraw);

const gridBounds = {
  maxV: Math.floor(page.gridWidth / step.diagonalX),
  maxU: Math.floor(page.gridHeight / step.vertical)
};

const gridPoints = [];
for (let v = 0; v <= gridBounds.maxV; v += 0.5) {
  for (let u = 0; u <= gridBounds.maxU; u += 0.5) {
    const canvasPoint = gridToCanvas([u, v]);
    gridPoints.push({
      grid: [u, v],
      x: canvasPoint.x,
      y: canvasPoint.y,
      isMajor: Number.isInteger(u) && Number.isInteger(v)
    });
  }
}

strokeWidthValue.textContent = `${strokeWidthInput.value} px`;
fontSizeValue.textContent = `${fontSizeInput.value} px`;
updateToolUI();

function gridToCanvas(point) {
  const [u, v] = point;
  return {
    x: page.marginLeft + v * step.diagonalX,
    y: page.marginTop + u * step.vertical + v * step.diagonalY
  };
}

function snapPoint(x, y) {
  if (!isInsideBoard(x, y)) {
    return null;
  }

  let bestPoint = null;
  let bestDistance = Infinity;

  for (const point of gridPoints) {
    if (point.y > page.marginTop + page.gridHeight) {
      continue;
    }
    const distance = Math.hypot(point.x - x, point.y - y);
    if (distance < bestDistance) {
      bestDistance = distance;
      bestPoint = point.grid;
    }
  }

  return bestDistance < 42 ? bestPoint : null;
}

function isInsideBoard(x, y) {
  return (
    x >= page.marginLeft - 8 &&
    x <= page.marginLeft + page.gridWidth + 8 &&
    y >= page.marginTop - 8 &&
    y <= page.marginTop + page.gridHeight + 8
  );
}

function nearestGridPoint(x, y) {
  if (!isInsideBoard(x, y)) {
    return null;
  }

  let bestPoint = null;
  let bestDistance = Infinity;

  for (const point of gridPoints) {
    if (point.y > page.marginTop + page.gridHeight) {
      continue;
    }
    const distance = Math.hypot(point.x - x, point.y - y);
    if (distance < bestDistance) {
      bestDistance = distance;
      bestPoint = point.grid;
    }
  }

  return bestPoint;
}

function getCanvasPoint(event) {
  const rect = canvas.getBoundingClientRect();
  return {
    x: ((event.clientX - rect.left) / rect.width) * canvas.width,
    y: ((event.clientY - rect.top) / rect.height) * canvas.height
  };
}

function getAnnotationById(id) {
  return state.annotations.find((annotation) => annotation.id === id) ?? null;
}

function getAnnotationMetrics(annotation) {
  const lines = (annotation.text || "").split("\n");
  const fontSize = Number(annotation.size);
  const lineHeight = fontSize * 1.2;

  ctx.save();
  ctx.font = `${fontSize}px ${annotation.font}`;
  let maxWidth = 0;
  for (const line of lines) {
    maxWidth = Math.max(maxWidth, ctx.measureText(line || " ").width);
  }
  ctx.restore();

  const paddingX = 8;
  const paddingY = 6;
  const boxWidth = maxWidth + paddingX * 2;
  const boxHeight = Math.max(lineHeight + paddingY * 2, lines.length * lineHeight + paddingY * 2);

  return { lines, fontSize, lineHeight, paddingX, paddingY, boxWidth, boxHeight };
}

function getAnnotationTransform(annotation) {
  const anchor =
    annotation.x != null && annotation.y != null ? { x: annotation.x, y: annotation.y } : gridToCanvas(annotation.point);
  const metrics = getAnnotationMetrics(annotation);
  const offsetX = 10;
  const offsetY = -8;
  return {
    ...metrics,
    centerX: anchor.x + offsetX + metrics.boxWidth / 2,
    centerY: anchor.y + offsetY + metrics.boxHeight / 2,
    rotationRadians: (Number(annotation.rotation) * Math.PI) / 180
  };
}

function positionAnnotationEditor(annotation) {
  const { centerX, centerY, fontSize } = getAnnotationTransform(annotation);
  const canvasRect = canvas.getBoundingClientRect();
  const scaleX = canvasRect.width / canvas.width;
  const scaleY = canvasRect.height / canvas.height;

  annotationEditor.style.left = `${centerX * scaleX}px`;
  annotationEditor.style.top = `${centerY * scaleY}px`;
  annotationEditor.style.transform = "translate(-50%, -50%)";
  annotationEditor.style.font = `${fontSize}px ${annotation.font}`;
  annotationEditor.style.color = annotation.color;
}

function resizeAnnotationEditor() {
  annotationEditor.style.height = "auto";
  annotationEditor.style.height = `${annotationEditor.scrollHeight}px`;
}

function normalizeDegrees(radians) {
  let degrees = (radians * 180) / Math.PI;
  while (degrees <= -180) degrees += 360;
  while (degrees > 180) degrees -= 360;
  return Math.round(degrees);
}

function drawBackground() {
  ctx.clearRect(0, 0, page.width, page.height);

  const paper = ctx.createLinearGradient(0, 0, page.width, page.height);
  paper.addColorStop(0, "#fffefb");
  paper.addColorStop(1, "#f9f6ee");
  ctx.fillStyle = paper;
  ctx.fillRect(0, 0, page.width, page.height);

  drawIsometricGrid();
  drawBranding();
}

function drawIsometricGrid() {
  ctx.save();
  ctx.strokeStyle = "#bcbcbc";
  ctx.lineWidth = 1;

  ctx.beginPath();
  ctx.rect(page.marginLeft, page.marginTop, page.gridWidth, page.gridHeight);
  ctx.clip();

  for (let x = 0; x <= page.gridWidth; x += step.diagonalX) {
    const px = page.marginLeft + x;
    ctx.beginPath();
    ctx.moveTo(px, page.marginTop);
    ctx.lineTo(px, page.marginTop + page.gridHeight);
    ctx.stroke();
  }

  const diagonalSpan = (page.gridHeight / step.diagonalY) * step.diagonalX;

  for (let startX = -diagonalSpan; startX <= page.gridWidth; startX += step.diagonalX) {
    ctx.beginPath();
    ctx.moveTo(page.marginLeft + startX, page.marginTop + page.gridHeight);
    ctx.lineTo(page.marginLeft + startX + diagonalSpan, page.marginTop);
    ctx.stroke();
  }

  for (let startX = 0; startX <= page.gridWidth + diagonalSpan; startX += step.diagonalX) {
    ctx.beginPath();
    ctx.moveTo(page.marginLeft + startX, page.marginTop + page.gridHeight);
    ctx.lineTo(page.marginLeft + startX - diagonalSpan, page.marginTop);
    ctx.stroke();
  }

  ctx.restore();

  ctx.lineWidth = 1.4;
  ctx.strokeStyle = "#9e9e9e";
  ctx.strokeRect(page.marginLeft, page.marginTop, page.gridWidth, page.gridHeight);
}

function drawBranding() {
  const x = 26;
  const y = page.marginTop + page.gridHeight + 8;
  const maxWidth = 330;
  const maxHeight = page.height - y - 18;

  ctx.save();
  if (logoImage.complete && logoImage.naturalWidth > 0) {
    const scale = Math.min(maxWidth / logoImage.naturalWidth, maxHeight / logoImage.naturalHeight);
    const drawWidth = logoImage.naturalWidth * scale;
    const drawHeight = logoImage.naturalHeight * scale;
    ctx.drawImage(logoImage, x, y, drawWidth, drawHeight);
  } else {
    ctx.fillStyle = "#1f3f96";
    ctx.fillRect(x, y, maxWidth, Math.max(48, maxHeight));
  }
  ctx.restore();
}

function drawSegments() {
  for (const segment of state.segments) {
    const a = gridToCanvas(segment.from);
    const b = gridToCanvas(segment.to);
    ctx.save();
    ctx.strokeStyle = segment.color;
    ctx.lineWidth = segment.width;
    ctx.lineCap = "round";
    ctx.lineJoin = "round";
    ctx.beginPath();
    ctx.moveTo(a.x, a.y);
    ctx.lineTo(b.x, b.y);
    ctx.stroke();
    ctx.restore();
  }
}

function drawAnnotations() {
  for (const annotation of state.annotations) {
    if (!annotation.text.trim() && annotation.id !== state.editingAnnotationId) {
      continue;
    }

    const { lines, fontSize, lineHeight, paddingX, paddingY, boxWidth, boxHeight, centerX, centerY, rotationRadians } =
      getAnnotationTransform(annotation);
    const isActive = state.activeAnnotationId === annotation.id;

    ctx.save();
    ctx.font = `${fontSize}px ${annotation.font}`;
    ctx.textBaseline = "top";
    ctx.translate(centerX, centerY);
    ctx.rotate(rotationRadians);

    ctx.fillStyle = "rgba(255, 254, 251, 0.92)";
    ctx.strokeStyle = isActive ? "rgba(31, 63, 150, 0.55)" : "rgba(19, 32, 57, 0.16)";
    ctx.lineWidth = isActive ? 1.5 : 1;
    ctx.beginPath();
    ctx.roundRect(-boxWidth / 2, -boxHeight / 2, boxWidth, boxHeight, 10);
    ctx.fill();
    ctx.stroke();

    ctx.fillStyle = annotation.color;
    lines.forEach((line, index) => {
      ctx.fillText(line, -boxWidth / 2 + paddingX, -boxHeight / 2 + paddingY + index * lineHeight);
    });

    if (isActive) {
      ctx.strokeStyle = "rgba(31, 63, 150, 0.45)";
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.moveTo(0, -boxHeight / 2);
      ctx.lineTo(0, -boxHeight / 2 - 18);
      ctx.stroke();

      ctx.fillStyle = "#ea580c";
      ctx.strokeStyle = "#132039";
      ctx.lineWidth = 2;
      ctx.beginPath();
      ctx.arc(0, -boxHeight / 2 - 18, 7, 0, Math.PI * 2);
      ctx.fill();
      ctx.stroke();
    }
    ctx.restore();
  }
}

function drawPendingGuide() {
  const anchor = state.pendingPoint ?? state.lineChainPoint;
  if (!anchor || !state.hoverPoint || !showGuidesInput.checked || state.tool !== "line") {
    return;
  }

  const a = gridToCanvas(anchor);
  const b = gridToCanvas(state.hoverPoint);
  ctx.save();
  ctx.strokeStyle = strokeColorInput.value;
  ctx.lineWidth = Number(strokeWidthInput.value);
  ctx.setLineDash([12, 10]);
  ctx.globalAlpha = 0.65;
  ctx.beginPath();
  ctx.moveTo(a.x, a.y);
  ctx.lineTo(b.x, b.y);
  ctx.stroke();
  ctx.restore();
}

function drawNodes() {
  if (!showNodesInput.checked) {
    return;
  }

  ctx.save();
  for (const point of gridPoints) {
    if (point.y > page.marginTop + page.gridHeight) {
      continue;
    }
    ctx.fillStyle = point.isMajor ? "rgba(19, 32, 57, 0.16)" : "rgba(19, 32, 57, 0.08)";
    ctx.beginPath();
    ctx.arc(point.x, point.y, point.isMajor ? 2.2 : 1.4, 0, Math.PI * 2);
    ctx.fill();
  }
  ctx.restore();

  if (!state.hoverPoint) {
    return;
  }

  const { x, y } = gridToCanvas(state.hoverPoint);
  ctx.save();
  ctx.fillStyle = "rgba(255, 255, 255, 0.96)";
  ctx.strokeStyle = "#1f3f96";
  ctx.lineWidth = 2;
  ctx.beginPath();
  ctx.arc(x, y, 6.5, 0, Math.PI * 2);
  ctx.fill();
  ctx.stroke();
  ctx.restore();
}

function redraw() {
  drawBackground();
  drawSegments();
  drawAnnotations();
  drawPendingGuide();
  drawNodes();
}

function addSegment(from, to) {
  state.segments.push({
    from,
    to,
    color: strokeColorInput.value,
    width: Number(strokeWidthInput.value)
  });
  state.history.push({ type: "segment" });
}

function createAnnotation(point) {
  const annotation = {
    id: state.nextAnnotationId,
    x: point.x,
    y: point.y,
    text: "",
    color: textColorInput.value,
    size: Number(fontSizeInput.value),
    font: fontFamilyInput.value,
    rotation: 0,
    isDraft: true
  };
  state.nextAnnotationId += 1;
  state.annotations.push(annotation);
  state.activeAnnotationId = annotation.id;
  return annotation;
}

function finishEditingAnnotation() {
  if (state.editingAnnotationId == null) {
    return;
  }

  const annotation = getAnnotationById(state.editingAnnotationId);
  if (!annotation) {
    state.editingAnnotationId = null;
    annotationEditor.hidden = true;
    annotationEditor.setAttribute("hidden", "");
    return;
  }

  annotation.text = annotationEditor.value;
  annotation.color = textColorInput.value;
  annotation.size = Number(fontSizeInput.value);
  annotation.font = fontFamilyInput.value;

  if (!annotation.text.trim()) {
    state.annotations = state.annotations.filter((item) => item.id !== annotation.id);
    if (state.activeAnnotationId === annotation.id) {
      state.activeAnnotationId = null;
    }
  } else if (annotation.isDraft) {
    annotation.isDraft = false;
    state.history.push({ type: "annotation", id: annotation.id });
  }

  state.editingAnnotationId = null;
  annotationEditor.hidden = true;
  annotationEditor.setAttribute("hidden", "");
  redraw();
}

function startEditingAnnotation(annotationId) {
  finishEditingAnnotation();
  const annotation = getAnnotationById(annotationId);
  if (!annotation) {
    return;
  }

  state.activeAnnotationId = annotation.id;
  state.editingAnnotationId = annotation.id;
  annotationEditor.hidden = false;
  annotationEditor.removeAttribute("hidden");
  annotationEditor.value = annotation.text;
  annotationEditor.placeholder = "Type annotation";
  positionAnnotationEditor(annotation);
  resizeAnnotationEditor();
  requestAnimationFrame(() => {
    annotationEditor.focus();
    annotationEditor.select();
  });
}

function hitTestAnnotation(x, y) {
  for (let index = state.annotations.length - 1; index >= 0; index -= 1) {
    const annotation = state.annotations[index];
    if (!annotation.text.trim() && annotation.id !== state.editingAnnotationId) {
      continue;
    }

    const { boxWidth, boxHeight, centerX, centerY, rotationRadians } = getAnnotationTransform(annotation);
    const dx = x - centerX;
    const dy = y - centerY;
    const cos = Math.cos(-rotationRadians);
    const sin = Math.sin(-rotationRadians);
    const localX = dx * cos - dy * sin;
    const localY = dx * sin + dy * cos;
    const handleY = -boxHeight / 2 - 18;

    if (Math.hypot(localX, localY - handleY) <= 10) {
      return { annotation, type: "handle" };
    }

    if (
      localX >= -boxWidth / 2 &&
      localX <= boxWidth / 2 &&
      localY >= -boxHeight / 2 &&
      localY <= boxHeight / 2
    ) {
      return { annotation, type: "body" };
    }
  }
  return null;
}

function updateToolUI() {
  const isLine = state.tool === "line";
  lineModeButton.classList.toggle("active", isLine);
  textModeButton.classList.toggle("active", !isLine);
  lineModeButton.classList.toggle("secondary", !isLine);
  textModeButton.classList.toggle("secondary", isLine);
  textControls.hidden = isLine;
}

canvas.addEventListener("pointermove", (event) => {
  const point = getCanvasPoint(event);
  state.hoverPoint = snapPoint(point.x, point.y);

  if (state.rotatingAnnotationId != null) {
    const annotation = getAnnotationById(state.rotatingAnnotationId);
    if (annotation) {
      const { centerX, centerY } = getAnnotationTransform(annotation);
      annotation.rotation = normalizeDegrees(Math.atan2(point.y - centerY, point.x - centerX) + Math.PI / 2);
      redraw();
    }
    canvas.style.cursor = "grabbing";
    return;
  }

  if (state.draggingAnnotationId != null) {
    const annotation = getAnnotationById(state.draggingAnnotationId);
    if (annotation) {
      annotation.x = point.x - state.dragOffset.x;
      annotation.y = point.y - state.dragOffset.y;
      redraw();
    }
    canvas.style.cursor = "grabbing";
    return;
  }

  if (state.pendingDragAnnotationId != null && state.pointerDownPoint) {
    const annotation = getAnnotationById(state.pendingDragAnnotationId);
    const moved = Math.hypot(
      point.x - state.pointerDownPoint.x,
      point.y - state.pointerDownPoint.y
    );
    if (annotation && moved > 8) {
      state.draggingAnnotationId = annotation.id;
      state.pendingDragAnnotationId = null;
      state.dragOffset = {
        x: state.pointerDownPoint.x - annotation.x,
        y: state.pointerDownPoint.y - annotation.y
      };
      canvas.style.cursor = "grabbing";
      return;
    }
  }

  if (state.tool === "text") {
    const hit = hitTestAnnotation(point.x, point.y);
    if (hit?.type === "handle") {
      canvas.style.cursor = "grab";
    } else if (hit?.type === "body") {
      canvas.style.cursor = "move";
    } else if (isInsideBoard(point.x, point.y)) {
      canvas.style.cursor = "text";
    } else {
      canvas.style.cursor = "default";
    }
  } else {
    canvas.style.cursor = state.hoverPoint ? "crosshair" : "default";
  }
  redraw();
});

canvas.addEventListener("pointerleave", () => {
  state.hoverPoint = null;
  if (state.rotatingAnnotationId == null) {
    canvas.style.cursor = "default";
  }
  redraw();
});

canvas.addEventListener("pointerdown", (event) => {
  const point = getCanvasPoint(event);

  if (state.tool === "text") {
    const hit = hitTestAnnotation(point.x, point.y);

    if (hit?.type === "handle") {
      finishEditingAnnotation();
      state.activeAnnotationId = hit.annotation.id;
      state.rotatingAnnotationId = hit.annotation.id;
      canvas.style.cursor = "grabbing";
      redraw();
      return;
    }

    if (hit?.type === "body") {
      finishEditingAnnotation();
      state.activeAnnotationId = hit.annotation.id;
      state.pendingDragAnnotationId = hit.annotation.id;
      state.pointerDownPoint = { x: point.x, y: point.y };
      redraw();
      return;
    }

    const insideBoard = isInsideBoard(point.x, point.y);
    finishEditingAnnotation();
    if (!insideBoard) {
      state.activeAnnotationId = null;
      redraw();
      return;
    }

    const annotation = createAnnotation({ x: point.x, y: point.y });
    startEditingAnnotation(annotation.id);
    redraw();
    return;
  }

  const snapped = state.hoverPoint ?? snapPoint(point.x, point.y);
  if (!snapped) {
    return;
  }

  if (!state.pendingPoint && !state.lineChainPoint) {
    state.pendingPoint = [...snapped];
  } else {
    const fromPoint = state.pendingPoint ?? state.lineChainPoint;
    addSegment(fromPoint, [...snapped]);
    state.pendingPoint = null;
    state.lineChainPoint = [...snapped];
  }
  redraw();
});

canvas.addEventListener("dblclick", (event) => {
  if (state.tool !== "line") {
    return;
  }

  const point = getCanvasPoint(event);
  const snapped = state.hoverPoint ?? snapPoint(point.x, point.y);
  if (!snapped || !state.lineChainPoint) {
    return;
  }

  const onActiveEndpoint =
    snapped[0] === state.lineChainPoint[0] &&
    snapped[1] === state.lineChainPoint[1];

  if (onActiveEndpoint) {
    state.pendingPoint = null;
    state.lineChainPoint = null;
    redraw();
  }
});

window.addEventListener("pointerup", () => {
  if (state.rotatingAnnotationId != null) {
    state.rotatingAnnotationId = null;
  }
  if (state.pendingDragAnnotationId != null) {
    const annotationId = state.pendingDragAnnotationId;
    state.pendingDragAnnotationId = null;
    state.pointerDownPoint = null;
    startEditingAnnotation(annotationId);
    return;
  }
  if (state.draggingAnnotationId != null) {
    state.draggingAnnotationId = null;
  }
  state.pointerDownPoint = null;
  redraw();
});

window.addEventListener("keydown", (event) => {
  if (event.key === "Escape") {
    finishEditingAnnotation();
    state.pendingPoint = null;
    state.lineChainPoint = null;
    state.activeAnnotationId = null;
    state.pendingDragAnnotationId = null;
    state.draggingAnnotationId = null;
    state.pointerDownPoint = null;
    redraw();
  }
});

strokeWidthInput.addEventListener("input", () => {
  strokeWidthValue.textContent = `${strokeWidthInput.value} px`;
  redraw();
});

fontSizeInput.addEventListener("input", () => {
  fontSizeValue.textContent = `${fontSizeInput.value} px`;
  const annotation = getAnnotationById(state.editingAnnotationId);
  if (annotation) {
    annotation.size = Number(fontSizeInput.value);
    positionAnnotationEditor(annotation);
    resizeAnnotationEditor();
    redraw();
  }
});

textColorInput.addEventListener("input", () => {
  const annotation = getAnnotationById(state.editingAnnotationId);
  if (annotation) {
    annotation.color = textColorInput.value;
    annotationEditor.style.color = annotation.color;
    redraw();
  }
});

fontFamilyInput.addEventListener("change", () => {
  const annotation = getAnnotationById(state.editingAnnotationId);
  if (annotation) {
    annotation.font = fontFamilyInput.value;
    positionAnnotationEditor(annotation);
    resizeAnnotationEditor();
    redraw();
  }
});

lineModeButton.addEventListener("click", () => {
  finishEditingAnnotation();
  state.tool = "line";
  updateToolUI();
  redraw();
});

textModeButton.addEventListener("click", () => {
  state.pendingPoint = null;
  state.lineChainPoint = null;
  state.tool = "text";
  updateToolUI();
  redraw();
});

showNodesInput.addEventListener("change", redraw);
showGuidesInput.addEventListener("change", redraw);

undoButton.addEventListener("click", () => {
  finishEditingAnnotation();
  const lastAction = state.history.pop();
  if (!lastAction) {
    return;
  }

  if (lastAction.type === "segment") {
    state.segments.pop();
  }

  if (lastAction.type === "annotation") {
    state.annotations = state.annotations.filter((annotation) => annotation.id !== lastAction.id);
    if (state.activeAnnotationId === lastAction.id) {
      state.activeAnnotationId = null;
    }
  }
  redraw();
});

clearButton.addEventListener("click", () => {
  finishEditingAnnotation();
  state.segments = [];
  state.annotations = [];
  state.history = [];
  state.pendingPoint = null;
  state.lineChainPoint = null;
  state.activeAnnotationId = null;
  state.draggingAnnotationId = null;
  state.pendingDragAnnotationId = null;
  state.pointerDownPoint = null;
  redraw();
});

exportButton.addEventListener("click", () => {
  finishEditingAnnotation();
  const link = document.createElement("a");
  link.href = canvas.toDataURL("image/png");
  link.download = "pmi-isometric-plot.png";
  link.click();
});

annotationEditor.addEventListener("input", () => {
  const annotation = getAnnotationById(state.editingAnnotationId);
  if (!annotation) {
    return;
  }

  annotation.text = annotationEditor.value;
  annotation.color = textColorInput.value;
  annotation.size = Number(fontSizeInput.value);
  annotation.font = fontFamilyInput.value;
  positionAnnotationEditor(annotation);
  resizeAnnotationEditor();
  redraw();
});

annotationEditor.addEventListener("blur", () => {
  finishEditingAnnotation();
});

annotationEditor.addEventListener("keydown", (event) => {
  if (event.key === "Escape") {
    event.preventDefault();
    finishEditingAnnotation();
  }
  if ((event.metaKey || event.ctrlKey) && event.key === "Enter") {
    event.preventDefault();
    annotationEditor.blur();
  }
});

window.addEventListener("resize", () => {
  const annotation = getAnnotationById(state.editingAnnotationId);
  if (annotation) {
    positionAnnotationEditor(annotation);
  }
});

redraw();
