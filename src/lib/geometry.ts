// Maps the shape names in a slide's XML (OOXML "prst" values) to the names PowerPoint's
// add-in API uses to create the same shape. The API can create a shape type but not report it,
// so the exported slide's XML is the only place Retro can learn which shape something is.

/** Every type PowerPoint.ShapeCollection.addGeometricShape accepts (from the Office.js typings). */
export const API_GEOMETRIES = new Set<string>([
  "LineInverse", "Triangle", "RightTriangle", "Rectangle", "Diamond", "Parallelogram", "Trapezoid",
  "NonIsoscelesTrapezoid", "Pentagon", "Hexagon", "Heptagon", "Octagon", "Decagon", "Dodecagon", "Star4", "Star5", "Star6",
  "Star7", "Star8", "Star10", "Star12", "Star16", "Star24", "Star32", "RoundRectangle", "Round1Rectangle",
  "Round2SameRectangle", "Round2DiagonalRectangle", "SnipRoundRectangle", "Snip1Rectangle",
  "Snip2SameRectangle", "Snip2DiagonalRectangle", "Plaque", "Ellipse", "Teardrop", "HomePlate", "Chevron", "PieWedge",
  "Pie", "BlockArc", "Donut", "NoSmoking", "RightArrow", "LeftArrow", "UpArrow", "DownArrow", "StripedRightArrow",
  "NotchedRightArrow", "BentUpArrow", "LeftRightArrow", "UpDownArrow", "LeftUpArrow", "LeftRightUpArrow", "QuadArrow",
  "LeftArrowCallout", "RightArrowCallout", "UpArrowCallout", "DownArrowCallout", "LeftRightArrowCallout",
  "UpDownArrowCallout", "QuadArrowCallout", "BentArrow", "UturnArrow", "CircularArrow", "LeftCircularArrow",
  "LeftRightCircularArrow", "CurvedRightArrow", "CurvedLeftArrow", "CurvedUpArrow", "CurvedDownArrow",
  "SwooshArrow", "Cube", "Can", "LightningBolt", "Heart", "Sun", "Moon", "SmileyFace", "IrregularSeal1", "IrregularSeal2",
  "FoldedCorner", "Bevel", "Frame", "HalfFrame", "Corner", "DiagonalStripe", "Chord", "Arc", "LeftBracket", "RightBracket",
  "LeftBrace", "RightBrace", "BracketPair", "BracePair", "Callout1", "Callout2", "Callout3", "AccentCallout1",
  "AccentCallout2", "AccentCallout3", "BorderCallout1", "BorderCallout2", "BorderCallout3", "AccentBorderCallout1",
  "AccentBorderCallout2", "AccentBorderCallout3", "WedgeRectCallout", "WedgeRRectCallout",
  "WedgeEllipseCallout", "CloudCallout", "Cloud", "Ribbon", "Ribbon2", "EllipseRibbon", "EllipseRibbon2",
  "LeftRightRibbon", "VerticalScroll", "HorizontalScroll", "Wave", "DoubleWave", "Plus", "FlowChartProcess",
  "FlowChartDecision", "FlowChartInputOutput", "FlowChartPredefinedProcess", "FlowChartInternalStorage",
  "FlowChartDocument", "FlowChartMultidocument", "FlowChartTerminator", "FlowChartPreparation",
  "FlowChartManualInput", "FlowChartManualOperation", "FlowChartConnector", "FlowChartPunchedCard",
  "FlowChartPunchedTape", "FlowChartSummingJunction", "FlowChartOr", "FlowChartCollate", "FlowChartSort",
  "FlowChartExtract", "FlowChartMerge", "FlowChartOfflineStorage", "FlowChartOnlineStorage",
  "FlowChartMagneticTape", "FlowChartMagneticDisk", "FlowChartMagneticDrum", "FlowChartDisplay",
  "FlowChartDelay", "FlowChartAlternateProcess", "FlowChartOffpageConnector", "ActionButtonBlank",
  "ActionButtonHome", "ActionButtonHelp", "ActionButtonInformation", "ActionButtonForwardNext",
  "ActionButtonBackPrevious", "ActionButtonEnd", "ActionButtonBeginning", "ActionButtonReturn",
  "ActionButtonDocument", "ActionButtonSound", "ActionButtonMovie", "Gear6", "Gear9", "Funnel", "MathPlus", "MathMinus",
  "MathMultiply", "MathDivide", "MathEqual", "MathNotEqual", "CornerTabs", "SquareTabs", "PlaqueTabs", "ChartX",
  "ChartStar", "ChartPlus",
]);

/** OOXML names that don't simply capitalise to the API name. */
const RENAMED: Record<string, string> = {
  lineInv: "LineInverse",
  rtTriangle: "RightTriangle",
  rect: "Rectangle",
  roundRect: "RoundRectangle",
  round1Rect: "Round1Rectangle",
  round2SameRect: "Round2SameRectangle",
  round2DiagRect: "Round2DiagonalRectangle",
  snipRoundRect: "SnipRoundRectangle",
  snip1Rect: "Snip1Rectangle",
  snip2SameRect: "Snip2SameRectangle",
  snip2DiagRect: "Snip2DiagonalRectangle",
  diagStripe: "DiagonalStripe",
  wedgeRoundRectCallout: "WedgeRRectCallout",
};

/** The API name for an OOXML preset shape, or undefined if the API can't create it. */
export function apiGeometry(prst: string): string | undefined {
  const name = RENAMED[prst] ?? prst.charAt(0).toUpperCase() + prst.slice(1);
  return API_GEOMETRIES.has(name) ? name : undefined;
}

/** Preset geometries that are lines rather than filled shapes. */
export const STRAIGHT_LINES = new Set(["line", "straightConnector1"]);
export const BENT_LINES = /^(bent|curved)Connector[2-5]$/;
