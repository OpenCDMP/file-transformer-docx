package org.opencdmp.filetransformer.docx.service.template.render;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.imageio.ImageIO;
import javax.imageio.ImageReader;
import javax.imageio.stream.ImageInputStream;
import java.io.ByteArrayInputStream;
import java.math.BigInteger;
import java.util.Iterator;
import java.util.Map;

/**
 * Low-level OOXML writing for one document: styled paragraphs and runs, hyperlinks, tables bound to
 * template table styles, pictures, lists and fields. Holds per-document state (list definitions,
 * figure counter), so one instance serves exactly one export.
 */
public final class DocxWriter {
    private static final Logger logger = LoggerFactory.getLogger(DocxWriter.class);

    private static final Map<String, Integer> PICTURE_TYPES = Map.of(
            "image/jpeg", Document.PICTURE_TYPE_JPEG,
            "image/png", Document.PICTURE_TYPE_PNG,
            "image/gif", Document.PICTURE_TYPE_GIF,
            "image/bmp", Document.PICTURE_TYPE_BMP,
            "image/tiff", Document.PICTURE_TYPE_TIFF);
    private static final String[] BULLETS = {"•", "◦", "–"};
    private static final String[] NUMBER_FORMATS = {"%1.", "%2.", "%3."};
    private static final STNumberFormat.Enum[] NUMBER_KINDS = {STNumberFormat.DECIMAL, STNumberFormat.LOWER_LETTER, STNumberFormat.LOWER_ROMAN};
    private static final int LIST_INDENT = 360;
    private static final int LIST_HANGING = 240;
    private static final java.util.regex.Pattern HEADER_FILL = java.util.regex.Pattern.compile(
            "tblStylePr[^>]*type=\"firstRow\".*?<(?:w:)?shd[^>]*fill=\"([0-9A-Fa-f]{6})\"", java.util.regex.Pattern.DOTALL);

    private final XWPFDocument document;
    private BigInteger bulletNumId;
    private BigInteger numberedAbstractId;
    private int figures;

    public DocxWriter(XWPFDocument document) {
        this.document = document;
    }

    public XWPFDocument document() {
        return this.document;
    }

    // ------------------------------------------------------------------ paragraphs and runs

    public static XWPFParagraph styled(XWPFParagraph paragraph, String style) {
        if (style != null) paragraph.setStyle(style);
        return paragraph;
    }

    public static XWPFRun text(XWPFParagraph paragraph, String text, String characterStyle) {
        XWPFRun run = paragraph.createRun();
        setText(run, text);
        if (characterStyle != null) characterStyle(run, characterStyle);
        return run;
    }

    /** Sets text, turning line breaks into Word line breaks. */
    public static void setText(XWPFRun run, String text) {
        if (text == null) return;
        String[] lines = text.replace("\r\n", "\n").split("\n", -1);
        for (int i = 0; i < lines.length; i++) {
            if (i > 0) run.addBreak();
            run.setText(lines[i], i);
        }
    }

    public static void characterStyle(XWPFRun run, String styleId) {
        CTRPr rpr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
        (rpr.sizeOfRStyleArray() > 0 ? rpr.getRStyleArray(0) : rpr.addNewRStyle()).setVal(styleId);
    }

    /** A hyperlink run; falls back to plain text when the address is not a valid URI. */
    public static XWPFRun link(XWPFParagraph paragraph, String url, String text) {
        try {
            java.net.URI uri = java.net.URI.create(url.trim().replace(" ", "%20"));
            if (uri.getScheme() == null) throw new IllegalArgumentException("relative uri");
            XWPFHyperlinkRun run = paragraph.createHyperlinkRun(uri.toASCIIString());
            run.setText(text);
            characterStyle(run, TemplateStyles.HYPERLINK);
            return run;
        } catch (Exception e) {
            return text(paragraph, text, null);
        }
    }

    public static void keepWithNext(XWPFParagraph paragraph) {
        CTPPr ppr = paragraph.getCTP().isSetPPr() ? paragraph.getCTP().getPPr() : paragraph.getCTP().addNewPPr();
        if (!ppr.isSetKeepNext()) ppr.addNewKeepNext();
    }

    public static void pageBreak(XWPFParagraph paragraph) {
        paragraph.createRun().addBreak(BreakType.PAGE);
    }

    // ------------------------------------------------------------------ lists

    /** Makes the paragraph an item of a bulleted list at the given nesting level. */
    public void bullet(XWPFParagraph paragraph, int level) {
        if (this.bulletNumId == null) this.bulletNumId = this.addList(true);
        this.numbering(paragraph, this.bulletNumId, level);
    }

    /** Starts a new numbered list (numbering restarts at 1) and returns its id for {@link #numbered}. */
    public BigInteger newNumberedList() {
        if (this.numberedAbstractId == null) {
            this.numberedAbstractId = this.addAbstractList(false);
        }
        XWPFNumbering numbering = this.document.createNumbering();
        BigInteger numId = numbering.addNum(this.numberedAbstractId);
        CTNum num = numbering.getNum(numId).getCTNum();
        for (int level = 0; level < NUMBER_FORMATS.length; level++) {
            CTNumLvl override = num.addNewLvlOverride();
            override.setIlvl(BigInteger.valueOf(level));
            override.addNewStartOverride().setVal(BigInteger.ONE);
        }
        return numId;
    }

    public void numbered(XWPFParagraph paragraph, BigInteger numId, int level) {
        this.numbering(paragraph, numId, level);
    }

    private void numbering(XWPFParagraph paragraph, BigInteger numId, int level) {
        paragraph.setNumID(numId);
        paragraph.setNumILvl(BigInteger.valueOf(Math.max(0, Math.min(2, level))));
    }

    private BigInteger addList(boolean bullets) {
        BigInteger abstractId = this.addAbstractList(bullets);
        return this.document.createNumbering().addNum(abstractId);
    }

    private BigInteger addAbstractList(boolean bullets) {
        XWPFNumbering numbering = this.document.createNumbering();
        BigInteger id = BigInteger.valueOf(900);
        while (numbering.getAbstractNum(id) != null) id = id.add(BigInteger.ONE);
        CTAbstractNum abstractNum = CTAbstractNum.Factory.newInstance();
        abstractNum.setAbstractNumId(id);
        for (int level = 0; level < 3; level++) {
            CTLvl lvl = abstractNum.addNewLvl();
            lvl.setIlvl(BigInteger.valueOf(level));
            lvl.addNewStart().setVal(BigInteger.ONE);
            lvl.addNewNumFmt().setVal(bullets ? STNumberFormat.BULLET : NUMBER_KINDS[level]);
            lvl.addNewLvlText().setVal(bullets ? BULLETS[level] : NUMBER_FORMATS[level]);
            lvl.addNewLvlJc().setVal(STJc.LEFT);
            CTInd ind = lvl.addNewPPr().addNewInd();
            ind.setLeft(BigInteger.valueOf((long) LIST_INDENT * (level + 1)));
            ind.setHanging(BigInteger.valueOf(LIST_HANGING));
        }
        return numbering.addAbstractNum(new XWPFAbstractNum(abstractNum, numbering));
    }

    // ------------------------------------------------------------------ tables

    /**
     * Binds a table to a template table style, full width, with the given relative column widths.
     * The header row repeats on every page and its cells are shaded explicitly with the style's
     * header fill, which keeps the look identical in Word and in the LibreOffice PDF conversion.
     */
    public void table(XWPFTable table, String styleId, double[] widths, boolean header) {
        CTTblPr tblPr = table.getCTTbl().getTblPr() != null ? table.getCTTbl().getTblPr() : table.getCTTbl().addNewTblPr();
        (tblPr.isSetTblStyle() ? tblPr.getTblStyle() : tblPr.addNewTblStyle()).setVal(styleId);
        CTTblWidth width = tblPr.isSetTblW() ? tblPr.getTblW() : tblPr.addNewTblW();
        width.setType(STTblWidth.PCT);
        width.setW(BigInteger.valueOf(5000));
        CTTblLook look = tblPr.isSetTblLook() ? tblPr.getTblLook() : tblPr.addNewTblLook();
        look.setVal(header ? new byte[]{0x04, (byte) 0xA0} : new byte[]{0x04, (byte) 0x80});
        CTTblLayoutType layout = tblPr.isSetTblLayout() ? tblPr.getTblLayout() : tblPr.addNewTblLayout();
        layout.setType(STTblLayoutType.FIXED);
        // remove the explicit borders POI adds so the table style decides
        if (tblPr.isSetTblBorders()) tblPr.unsetTblBorders();

        int contentWidth = this.contentWidth();
        CTTblGrid grid = table.getCTTbl().getTblGrid() != null ? table.getCTTbl().getTblGrid() : table.getCTTbl().addNewTblGrid();
        while (grid.sizeOfGridColArray() > 0) grid.removeGridCol(0);
        double total = 0;
        for (double w : widths) total += w;
        for (double w : widths) grid.addNewGridCol().setW(BigInteger.valueOf(Math.round(contentWidth * w / total)));

        for (int r = 0; r < table.getRows().size(); r++) {
            XWPFTableRow row = table.getRow(r);
            for (int c = 0; c < row.getTableCells().size() && c < widths.length; c++) {
                XWPFTableCell cell = row.getCell(c);
                CTTcPr tcPr = cell.getCTTc().isSetTcPr() ? cell.getCTTc().getTcPr() : cell.getCTTc().addNewTcPr();
                CTTblWidth cellWidth = tcPr.isSetTcW() ? tcPr.getTcW() : tcPr.addNewTcW();
                cellWidth.setType(STTblWidth.DXA);
                cellWidth.setW(BigInteger.valueOf(Math.round(contentWidth * widths[c] / total)));
            }
        }
        if (header && !table.getRows().isEmpty()) {
            XWPFTableRow headerRow = table.getRow(0);
            headerRow.setRepeatHeader(true);
            headerRow.setCantSplitRow(true);
            String fill = this.headerFill(styleId);
            if (fill != null) headerRow.getTableCells().forEach(cell -> cell.setColor(fill));
        }
    }

    /** The first-row shading of a table style, if the style defines one (read from XML: the lite schemas lack tblStylePr). */
    private String headerFill(String styleId) {
        try {
            XWPFStyle style = this.document.getStyles() != null ? this.document.getStyles().getStyle(styleId) : null;
            if (style == null) return null;
            java.util.regex.Matcher matcher = HEADER_FILL.matcher(style.getCTStyle().xmlText());
            return matcher.find() ? matcher.group(1) : null;
        } catch (Exception e) {
            logger.debug("no header fill for table style {}", styleId, e);
            return null;
        }
    }

    /** Page width minus margins, in twips (dxa), of the document's last section. */
    public int contentWidth() {
        try {
            CTSectPr sectPr = this.document.getDocument().getBody().getSectPr();
            int width = ((BigInteger) sectPr.getPgSz().getW()).intValue();
            int left = ((BigInteger) sectPr.getPgMar().getLeft()).intValue();
            int right = ((BigInteger) sectPr.getPgMar().getRight()).intValue();
            return Math.max(2000, width - left - right);
        } catch (Exception e) {
            return 9638; // A4 with 2 cm margins
        }
    }

    // ------------------------------------------------------------------ pictures and fields

    /** Adds a picture scaled to fit the content width (and a third of the page height at most), then a numbered caption. */
    public boolean picture(Sink sink, byte[] data, String mimeType, String filename, String captionLabel) {
        Integer type = PICTURE_TYPES.get(mimeType);
        if (type == null || data == null) return false;
        try (ImageInputStream input = ImageIO.createImageInputStream(new ByteArrayInputStream(data))) {
            Iterator<ImageReader> readers = ImageIO.getImageReaders(input);
            if (!readers.hasNext()) return false;
            ImageReader reader = readers.next();
            reader.setInput(input);
            double widthPt = reader.getWidth(0) * 0.75;
            double heightPt = reader.getHeight(0) * 0.75;
            double maxWidthPt = this.contentWidth() / 20.0 * (sink.inTable() ? 0.45 : 1.0);
            double maxHeightPt = 300;
            double scale = Math.min(1.0, Math.min(maxWidthPt / widthPt, maxHeightPt / heightPt));

            XWPFParagraph paragraph = sink.paragraph();
            paragraph.setAlignment(sink.inTable() ? ParagraphAlignment.LEFT : ParagraphAlignment.CENTER);
            keepWithNext(paragraph);
            paragraph.createRun().addPicture(new ByteArrayInputStream(data), type, filename != null ? filename : "image",
                    Units.toEMU(widthPt * scale), Units.toEMU(heightPt * scale));
            if (!sink.inTable()) {
                this.figures++;
                XWPFParagraph caption = styled(sink.paragraph(), TemplateStyles.CAPTION);
                caption.setAlignment(ParagraphAlignment.CENTER);
                text(caption, captionLabel + " " + this.figures + (filename != null && !filename.isBlank() ? "  ·  " + filename : ""), null);
            }
            return true;
        } catch (Exception e) {
            logger.error("could not add picture {}", filename, e);
            return false;
        }
    }

    /**
     * Appends a complex field (e.g. TOC, PAGE). The runs added by {@code result} become the field's
     * current result, shown until Word or the PDF conversion updates the field.
     */
    public static void field(XWPFParagraph paragraph, String instruction, Runnable result) {
        paragraph.createRun().getCTR().addNewFldChar().setFldCharType(STFldCharType.BEGIN);
        CTText instr = paragraph.createRun().getCTR().addNewInstrText();
        instr.setStringValue(instruction);
        instr.setSpace(org.apache.xmlbeans.impl.xb.xmlschema.SpaceAttribute.Space.PRESERVE);
        paragraph.createRun().getCTR().addNewFldChar().setFldCharType(STFldCharType.SEPARATE);
        if (result != null) result.run();
        paragraph.createRun().getCTR().addNewFldChar().setFldCharType(STFldCharType.END);
    }
}
