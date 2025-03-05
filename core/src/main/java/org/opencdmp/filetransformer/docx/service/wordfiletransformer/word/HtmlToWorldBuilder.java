package org.opencdmp.filetransformer.docx.service.wordfiletransformer.word;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Node;
import org.jsoup.nodes.TextNode;
import org.jsoup.select.NodeTraversor;
import org.jsoup.select.NodeVisitor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.math.BigInteger;
import java.net.URI;
import java.net.URL;
import java.util.function.Predicate;
import java.util.*;

public class HtmlToWorldBuilder implements NodeVisitor {

    private static final Logger log = LoggerFactory.getLogger(HtmlToWorldBuilder.class);
    private final Map<String, Boolean> properties = new LinkedHashMap<>();
    private XWPFParagraph paragraph;
    private XWPFRun run;
    private Boolean dumpRun;
    private final float indentation;
    private Boolean isIdentationUsed;
    private XWPFNumbering numbering;
    private final Queue<BigInteger> abstractNumId;
    private BigInteger numberingLevel;
    private XmlCursor cursor;

    public static HtmlToWorldBuilder convertInTable(XWPFTableCell document, Document htmlDocument, float indentation) {
        XWPFParagraph paragraph = document.addParagraph();
        paragraph.setIndentFromLeft(Math.round(400 * indentation));
        HtmlToWorldBuilder htmlToWorldBuilder = new HtmlToWorldBuilder(paragraph, indentation, null);
        NodeTraversor.traverse(htmlToWorldBuilder, htmlDocument);
        return htmlToWorldBuilder;
    }

    public static HtmlToWorldBuilder convert(XWPFDocument document, Document htmlDocument, float indentation) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setIndentFromLeft(Math.round(400 * indentation));
        HtmlToWorldBuilder htmlToWorldBuilder = new HtmlToWorldBuilder(paragraph, indentation, null);
        NodeTraversor.traverse(htmlToWorldBuilder, htmlDocument);
        return htmlToWorldBuilder;
    }

    public HtmlToWorldBuilder(XWPFParagraph paragraph, float indentation, XmlCursor cursor) {
        this.paragraph = paragraph;
        this.run = this.paragraph.createRun();
        this.dumpRun = false;
        this.indentation = indentation;
        this.isIdentationUsed = false;
        this.run.setFontSize(11);
        this.abstractNumId = new ArrayDeque<>();
        this.numberingLevel = BigInteger.valueOf(-1);
        this.setDefaultIndentation();
        this.cursor  = cursor;
    }

    @Override
    public void head(Node node, int i) {

        if (!node.outerHtml().contains("br")) {
            String htmlToPlainText = Jsoup.parse(node.outerHtml()).text();
            if (htmlToPlainText.trim().isEmpty()) return;
        }

        String name = node.nodeName();
        if (name.equals("#text")) {
            String text = ((TextNode)node).text();
            this.run.setText(text);
            this.dumpRun = true;
        } else {
            properties.put(name, true);
        }
        if (dumpRun) {
            this.run = this.paragraph.createRun();
            this.run.setFontSize(11);
            this.dumpRun = false;
        }
        parseProperties(node);
        properties.clear();
    }

    private void parseProperties(Node node) {
        properties.entrySet().forEach(stringBooleanEntry -> {
            this.run.setFontSize(11);
            switch (stringBooleanEntry.getKey()) {
                case "i" :
                case "em":
                    this.run.setItalic(stringBooleanEntry.getValue());
                    break;
                case "b":
                case "strong":
                    this.run.setBold(stringBooleanEntry.getValue());
                    break;
                case "u":
                case "ins":
                    this.run.setUnderline(stringBooleanEntry.getValue() ? UnderlinePatterns.SINGLE : UnderlinePatterns.NONE);
                    break;
                case "small":
                    this.run.setFontSize(stringBooleanEntry.getValue() ? 8 : 11);
                    break;
                case "del":
                case "strike":
                case "strikethrough":
                case "s":
                    this.run.setStrikeThrough(stringBooleanEntry.getValue());
                    break;
                case "mark":
                    this.run.setTextHighlightColor(stringBooleanEntry.getValue() ? STHighlightColor.YELLOW.toString() : STHighlightColor.NONE.toString());
                    break;
                case "sub":
                    this.run.setSubscript(stringBooleanEntry.getValue() ? VerticalAlign.SUBSCRIPT : VerticalAlign.BASELINE);
                    break;
                case "sup":
                    this.run.setSubscript(stringBooleanEntry.getValue() ? VerticalAlign.SUPERSCRIPT : VerticalAlign.BASELINE);
                    break;
                case "div":
                    if (node.childNodeSize() > 0) {
                        Predicate<Node> hasChildrenPredicate = child -> {
                            if (child.childNodeSize() == 0) return false;
                            return child.childNodes().stream().anyMatch(nch -> nch.childNodeSize() > 0);
                        };
                        boolean hasNodeNestedChildren = node.childNodes().stream().anyMatch(hasChildrenPredicate);
                        if (hasNodeNestedChildren) {
                            break;
                        }
                    }
                case "p":
                    if(this.cursor != null) {
                        this.paragraph = this.paragraph.getDocument().insertNewParagraph(this.cursor);
                        this.cursor = this.paragraph.getCTP().newCursor();
                        this.cursor.toNextSibling();
                    } else {
                        this.paragraph = this.paragraph.getDocument().createParagraph();
                    }
                    this.paragraph.setSpacingBefore(0);
                    this.paragraph.setSpacingAfter(0);
                    this.run = this.paragraph.createRun();
                    this.isIdentationUsed = false;
                    this.setDefaultIndentation();
                    if (stringBooleanEntry.getValue()) {
                        if (node.hasAttr("align")) {
                            String alignment = node.attr("align");
                            if(alignment.toUpperCase(Locale.ROOT).equals("JUSTIFY")) {
                                alignment = "both";
                            }
                            this.paragraph.setAlignment(ParagraphAlignment.valueOf(alignment.toUpperCase(Locale.ROOT)));
                        }
                    }
                    break;
                case "blockquote":
                    if(this.cursor != null) {
                        this.paragraph = this.paragraph.getDocument().insertNewParagraph(this.cursor);
                        this.cursor = this.paragraph.getCTP().newCursor();
                        cursor.toNextSibling();
                    } else {
                        this.paragraph = this.paragraph.getDocument().createParagraph();
                    }
                    this.run = this.paragraph.createRun();
                    if (stringBooleanEntry.getValue()) {
                        this.paragraph.setIndentationLeft(400);
                    } else {
                        this.isIdentationUsed = false;
                        this.setDefaultIndentation();
                    }
                    break;
                case "ul":
                    if (stringBooleanEntry.getValue()) {
                        createNumbering(STNumberFormat.BULLET);
                    } else {
                        if(this.cursor != null) {
                            this.paragraph = this.paragraph.getDocument().insertNewParagraph(this.cursor);
                            this.cursor = this.paragraph.getCTP().newCursor();
                            cursor.toNextSibling();
                        } else {
                            this.paragraph = this.paragraph.getDocument().createParagraph();
                        }
                        this.paragraph.setSpacingBefore(0);
                        this.paragraph.setSpacingAfter(0);
                        this.run = this.paragraph.createRun();
                        this.isIdentationUsed = false;
                        this.setDefaultIndentation();
                        this.numberingLevel = this.numberingLevel.subtract(BigInteger.ONE);
                        ((ArrayDeque)this.abstractNumId).removeLast();
                    }
                    break;
                case "ol":
                    if (stringBooleanEntry.getValue()) {
                        createNumbering(STNumberFormat.DECIMAL);
                    } else {
                        if(this.cursor != null) {
                            this.paragraph = this.paragraph.getDocument().insertNewParagraph(this.cursor);
                            this.cursor = this.paragraph.getCTP().newCursor();
                            cursor.toNextSibling();
                        } else {
                            this.paragraph = this.paragraph.getDocument().createParagraph();
                        }
                        this.paragraph.setSpacingBefore(0);
                        this.paragraph.setSpacingAfter(0);
                        this.run = this.paragraph.createRun();
                        this.isIdentationUsed = false;
                        this.setDefaultIndentation();
                        this.numberingLevel = this.numberingLevel.subtract(BigInteger.ONE);
                        ((ArrayDeque)this.abstractNumId).removeLast();
                    }
                    break;
                case "li":
                    if (stringBooleanEntry.getValue()) {
                        if(this.cursor != null) {
                            this.paragraph = this.paragraph.getDocument().insertNewParagraph(this.cursor);
                            this.cursor = this.paragraph.getCTP().newCursor();
                            cursor.toNextSibling();
                        } else {
                            this.paragraph = this.paragraph.getDocument().createParagraph();
                        }
                        //                            this.paragraph.setIndentationLeft(Math.round(indentation * 720) * (numberingLevel.intValue() + 1));
                        this.paragraph.setIndentFromLeft(Math.round(numberingLevel.intValue() * 400 + this.indentation*400));
                        this.run = this.paragraph.createRun();
                        this.paragraph.setNumID(((ArrayDeque<BigInteger>)abstractNumId).getLast().add(BigInteger.ONE)); // sets the list-element numId that should match the numId of the corresponding list
                    }
                    break;
                case "font":
                    if (stringBooleanEntry.getValue()) {
                        if (node.hasAttr("color")) {
                            this.run.setColor(node.attr("color").substring(1));
                        }
                    } else {
                        this.run.setColor("000000");
                    }
                    break;
                case "a":
                    if (stringBooleanEntry.getValue()) {
                        if (node.hasAttr("href")) {
                            XWPFHyperlinkRun xwpfHyperlinkRun = createHyperLinkRun(node.attr("href"));
                            if (xwpfHyperlinkRun != null) {
                                this.run = xwpfHyperlinkRun;
                                this.run.setColor("0000FF");
                                this.run.setUnderline(UnderlinePatterns.SINGLE);
                            } else {
                                this.run.setText(node.attr("href") + " ");
                            }
                        }
                    } else {
                        this.run = paragraph.createRun();
                    }
                    break;
                case "br":
                    if (stringBooleanEntry.getValue()) {
                        this.run.addBreak();
                    }
                    break;
                case "h1":
                    this.run.setFontSize(24);
                    break;
                case "h2":
                    this.run.setFontSize(20);
                    break;
                case "h3":
                    this.run.setFontSize(16);
                    break;
                case "h4":
                    this.run.setFontSize(14);
                    this.run.setBold(stringBooleanEntry.getValue());
                    break;
                case "h5":
                    this.run.setFontSize(14);
                    break;
                case "h6":
                    this.run.setFontSize(11);
                    this.run.setBold(stringBooleanEntry.getValue());
                    this.run.setCapitalized(stringBooleanEntry.getValue());
                    break;
            }
        });
    }

    @Override
    public void tail(Node node, int i) {
        if (!node.outerHtml().contains("br")) {
            String htmlToPlainText = Jsoup.parse(node.outerHtml()).text();
            if (htmlToPlainText.trim().isEmpty()) return;
        }

        String name = node.nodeName();
        properties.put(name, false);
        parseProperties(node);
        properties.clear();
    }

    //GK: This function creates one numbering.xml for the word document and adds a specific format.
    //It imitates the numbering.xml that is usually generated by word editors like LibreOffice
    private void createNumbering(STNumberFormat.Enum format) {
        CTAbstractNum ctAbstractNum = CTAbstractNum.Factory.newInstance();
        if (this.numbering == null) this.numbering = this.paragraph.getDocument().createNumbering();
        BigInteger tempNumId = BigInteger.ONE;
        boolean found = false;
        while (!found) {
            Object o = numbering.getAbstractNum(tempNumId);
            found = (o == null);
            if (!found) tempNumId = tempNumId.add(BigInteger.ONE);
        }
        ctAbstractNum.setAbstractNumId(tempNumId);
        CTLvl ctLvl = ctAbstractNum.addNewLvl();
        this.numberingLevel = numberingLevel.add(BigInteger.ONE);
        ctLvl.setIlvl(numberingLevel);
        ctLvl.addNewNumFmt().setVal(format);
        ctLvl.addNewStart().setVal(BigInteger.ONE);
        if (format == STNumberFormat.BULLET) {
            ctLvl.addNewLvlJc().setVal(STJc.LEFT);
            ctLvl.addNewLvlText().setVal("\u2022");
            ctLvl.addNewRPr(); //Set the Symbol font
            CTFonts f = ctLvl.getRPr().addNewRFonts();
            f.setAscii("Symbol");
            f.setHAnsi("Symbol");
            f.setCs("Symbol");
            f.setHint(STHint.DEFAULT);
        } else {
            ctLvl.addNewLvlText().setVal("%1.");
        }
        XWPFAbstractNum xwpfAbstractNum = new XWPFAbstractNum(ctAbstractNum);

        this.abstractNumId.add(this.numbering.addAbstractNum(xwpfAbstractNum));
        this.numbering.addNum(((ArrayDeque<BigInteger>)abstractNumId).getLast());
    }

    private XWPFHyperlinkRun createHyperLinkRun(String uri) {
        URI fixedUri = null;
        try {
            fixedUri = URI.create(uri);
        } catch (Exception e){
            log.warn("cannot create uri: " + uri);
        }
        if (fixedUri == null){
            try {
                fixedUri = URI.create(uri.replace(" ", "%20"));
            } catch (Exception e){
                log.warn("cannot create uri: " + uri);
            }
        }
        if (fixedUri == null) return null;
        
        String rId = this.paragraph.getDocument().getPackagePart().addExternalRelationship(fixedUri.toASCIIString(), XWPFRelation.HYPERLINK.getRelation()).getId();

        CTHyperlink cthyperLink=paragraph.getCTP().addNewHyperlink();
        cthyperLink.setId(rId);
        cthyperLink.addNewR();

        return new XWPFHyperlinkRun(
                cthyperLink,
                cthyperLink.getRArray(0),
                paragraph
        );
    }

    private void setDefaultIndentation() {
        if (!isIdentationUsed) {
//            this.paragraph.setIndentationLeft(Math.round(indentation * 720.0F));
            this.paragraph.setIndentFromLeft(Math.round(indentation * 400));
            this.isIdentationUsed = true;
        }
    }

    public XWPFParagraph getParagraph() {
        return paragraph;
    }
}
