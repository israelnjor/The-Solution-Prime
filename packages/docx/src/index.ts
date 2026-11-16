export { readDocx, type ReadResult } from "./reader.js";
export { writeDocx } from "./writer.js";
export { readParagraphProps, readRunProps, writeParagraphProps, writeRunProps } from "./props.js";
export { type DocumentProperties, readProperties, withProperties } from "./properties.js";
export { isCfb, readCfb, writeCfb } from "./cfb.js";
export { decryptDocx, encryptDocx, isEncrypted, WrongPasswordError } from "./encryption.js";
export { anchorXml, avLstXml, DEFAULT_THEME, lineXml, withShapeAdjust, type NewShape, newShapeXml, readShapeStyle, readShapeTransform, themeColors, withShapeFill, withShapeLine, withShapeText, withShapeTransform, withStacking, withTextAnchor } from "./shapes.js";
export { addChart, type ChartData, chartWorkbook, chartXml, NEW_CHART_COLORS, updateChart } from "./charts.js";
