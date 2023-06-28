package com.example.old_template_transform_web;

import com.aspose.words.*;
import com.aspose.words.Font;
import com.aspose.words.Shape;

import java.awt.*;
import java.io.*;
import java.lang.reflect.Modifier;
import java.util.ArrayList;

public class TransformUtil2 {

    /**
     * 将域（被替换的值）转换为双括号形式，转换复选框
     *
     * @return
     */
    public static byte[] procCustomDocument(InputStream inputStream) {
        byte[] bytes = null;
        try {
            Class<?> aClass = Class.forName("com.aspose.words.zzXyu");
            java.lang.reflect.Field zzYAC = aClass.getDeclaredField("zzZXG");
            zzYAC.setAccessible(true);

            java.lang.reflect.Field modifiersField = zzYAC.getClass().getDeclaredField("modifiers");
            modifiersField.setAccessible(true);
            modifiersField.setInt(zzYAC, zzYAC.getModifiers() & ~Modifier.FINAL);
            zzYAC.set(null, new byte[]{76, 73, 67, 69, 78, 83, 69, 68});

            ByteArrayOutputStream os = new ByteArrayOutputStream();
            Document doc = new Document(inputStream);

            // 处理自定义域
            CustomDocumentProperties documentProperties = doc.getCustomDocumentProperties();
            for (DocumentProperty obj : documentProperties) {
                String value = "{{" + obj.getName().replace("r_", "") + "}}";
                // 按照马宇泓的要求替换值
                if (value.equals("{{jj_fkfs}}")) {
                    value = "{{jj_jjfs}}";
                } else if (value.equals("{{jj_fkfsnew}}")) {
                    value = "{{jj_fkfs}}";
                } else if (value.equals("{{jgjg}}")) {
                    value = "{{jg_yh}}";
                } else if (value.equals("{{jgzhmc}}")) {
                    value = "{{jg_hm}}";
                } else if (value.equals("{{jgyhzh}}")) {
                    value = "{{jg_zh}}";
                } else if (value.equals("{{cm_fddbr}}")) {
                    value = "{{cm_bossname}}";
                } else if (value.equals("{{cm_tel}}")) {
                    value = "{{cm_lxdh}}";
                }
                obj.setValue(value);
            }
            doc.updateFields();
            doc.removeMacros();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // 处理控件
            NodeCollection nodeCollection = doc.getChildNodes(NodeType.SHAPE, true);
            ArrayList<Shape> moveList = new ArrayList<>();

            for (int i = 0; i < nodeCollection.getCount(); i++) {
                Shape shape = (Shape) nodeCollection.get(i);

                if (shape.getOleFormat() != null && shape.getOleFormat().getProgId().indexOf("Forms.CheckBox") >= 0) {
                    moveList.add(shape);
                    Forms2OleControl forms2OleControl = (Forms2OleControl) shape.getOleFormat().getOleControl();

                    StructuredDocumentTag tag = new StructuredDocumentTag(doc, SdtType.CHECKBOX, MarkupLevel.INLINE);
                    if (forms2OleControl.getValue().equals("1")) {
                        tag.setChecked(true);
                    }
                    tag.setCheckedSymbol(8730, "宋体");

                    builder.moveTo(shape);
                    builder.insertNode(tag);
                    builder.write(forms2OleControl.getCaption());
                }

                // 删掉老模板第一页上面的两个条形码图片，wps需要缩放至100%以下才能看到
                if (shape.getOleFormat() != null && shape.getOleFormat().getProgId().indexOf("BARCODE.BarCodeCtrl") >= 0) {
                    moveList.add(shape);
                }
            }

            for (int i = 0; i < moveList.size(); i++) {
                moveList.get(i).remove();
            }

            removeHeaderFooter(doc);

            addHeader(doc);

            addFooter(doc);

            doc.save(os, SaveFormat.DOCX);
            bytes = os.toByteArray();
            os.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return bytes;
    }

    static void removeHeaderFooter(Document doc) {
        for (Section section : doc.getSections()) {
            HeaderFooter header;
            HeaderFooter footer;
            header = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_FIRST);
            footer = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_FIRST);
            if (header != null)
                header.remove();
            if (footer != null)
                footer.remove();

            header = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_PRIMARY);
            footer = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_PRIMARY);
            if (header != null)
                header.remove();
            if (footer != null)
                footer.remove();

            header = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_EVEN);
            footer = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_EVEN);
            if (header != null)
                header.remove();
            if (footer != null)
                footer.remove();
        }
    }

    static void addHeader(Document doc) throws Exception {
        // 文档构建工具类，可对当前加入的模板进行编辑、新增等部分功能。
        DocumentBuilder builder = new DocumentBuilder(doc);
        // 将光标移动到页眉位置
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        // 靠右对齐
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.RIGHT);
        Paragraph paragraph = builder.insertParagraph();
        Run run = new Run(doc, "{{qrCode_ewm}}");
        paragraph.appendChild(run);

        // 添加页眉线
        Border borderHeader = builder.getParagraphFormat().getBorders().getBottom();
        borderHeader.setShadow(true);
        borderHeader.setDistanceFromText(2);
        borderHeader.setLineStyle(LineStyle.SINGLE);
    }

    static void addFooter(Document doc) throws Exception {
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < doc.getSections().getCount(); i++) {
            builder.moveToSection(i);
            builder.moveToHeaderFooter(HeaderFooterType.FOOTER_PRIMARY);
            builder.getCurrentParagraph().getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
            builder.getFont().setName("宋体");
            builder.getFont().setSize(9.0);
            builder.insertField(FieldType.FIELD_PAGE, true);
            builder.getPageSetup().setFooterDistance(35.0);
        }
    }

    /**
     * 将可编辑域转换为格式文本内容控件
     *
     * @param bytes
     * @param path
     */
    public static void procEditAbleRange(byte[] bytes, String path) {
        try {
            Class<?> aClass = Class.forName("com.aspose.words.zzXyu");
            java.lang.reflect.Field zzYAC = aClass.getDeclaredField("zzZXG");
            zzYAC.setAccessible(true);
            java.lang.reflect.Field modifiersField = zzYAC.getClass().getDeclaredField("modifiers");
            modifiersField.setAccessible(true);
            modifiersField.setInt(zzYAC, zzYAC.getModifiers() & ~Modifier.FINAL);
            zzYAC.set(null, new byte[]{76, 73, 67, 69, 78, 83, 69, 68});

            File file = new File(path);
            FileOutputStream os = new FileOutputStream(file);

            ByteArrayInputStream inputStream = new ByteArrayInputStream(bytes);
            Document doc = new Document(inputStream);
            DocumentBuilder builder = new DocumentBuilder(doc);
            NodeCollection startNodeCollection = doc.getChildNodes(NodeType.EDITABLE_RANGE_START, true);

            int count = startNodeCollection.getCount();
            for (int i = 0; i < count; i++) {
                EditableRangeStart rangeStart = (EditableRangeStart) startNodeCollection.get(0);
                EditableRangeEnd rangeEnd = rangeStart.getEditableRange().getEditableRangeEnd();
                ArrayList extractedNodes;
                if (rangeStart.getNextSibling() == null) {
                    extractedNodes = ExtractContentHelper.extractContent(rangeStart, rangeEnd, true);
                } else {
                    extractedNodes = ExtractContentHelper.extractContent(rangeStart.getNextSibling(), rangeEnd, true);
                }

                if (extractedNodes.size() > 1) {
                    StructuredDocumentTagRangeStart startTag = new StructuredDocumentTagRangeStart(doc, SdtType.RICH_TEXT);
                    startTag.isShowingPlaceholderText(false);
                    StructuredDocumentTagRangeEnd endTag = new StructuredDocumentTagRangeEnd(doc, startTag.getId());

                    builder.moveTo(rangeStart);
                    builder.getCurrentSection().getBody().insertBefore(startTag, builder.getCurrentParagraph());

                    if ((rangeEnd.getPreviousSibling() != null) && (rangeEnd.getPreviousSibling().getText().equals("\f"))
                            && (rangeEnd.getNextSibling() != null)) {
                        rangeEnd.getPreviousSibling().remove();
                        builder.moveTo(rangeEnd.getNextSibling());
                        Paragraph paragraph = builder.insertParagraph();
                        Run run = new Run(doc,"\f");
                        paragraph.prependChild(run);
                        builder.moveTo(rangeEnd);
                        builder.getCurrentSection().getBody().insertAfter(endTag, builder.getCurrentParagraph());
                    } else if ((rangeEnd.getNextSibling() != null) && (rangeEnd.getNextSibling().getText().equals("\f"))
                            && (rangeEnd.getPreviousSibling() != null)) {
                        rangeEnd.getNextSibling().remove();
                        builder.moveTo(rangeEnd.getNextSibling());
                        Paragraph paragraph = builder.insertParagraph();
                        Run run = new Run(doc,"\f");
                        paragraph.prependChild(run);
                        builder.moveTo(rangeEnd);
                        builder.getCurrentSection().getBody().insertAfter(endTag, builder.getCurrentParagraph());
                    } else {
                        builder.moveTo(rangeEnd);
                        Node currentNode = rangeEnd;
                        Paragraph currentPara = builder.getCurrentParagraph();

                        boolean flag = false;
                        Table table = null;
                        while ((currentNode.getNodeType() == NodeType.EDITABLE_RANGE_END) ||
                                ((currentNode.getNodeType() == NodeType.RUN) && (currentNode.getText().trim().isEmpty()))) {
                            currentNode = currentNode.getPreviousSibling();
                            while (currentNode == null) {
                                if (currentPara.getPreviousSibling() == null) {
                                    flag = true;
                                    break;
                                }
                                if (currentPara.getPreviousSibling().getNodeType() == NodeType.TABLE) {
                                    flag = true;
                                    table = (Table) currentPara.getPreviousSibling();
                                    break;
                                } else if (currentPara.getPreviousSibling().getNodeType() == NodeType.PARAGRAPH) {
                                    currentPara = (Paragraph) currentPara.getPreviousSibling();
                                    currentNode = currentPara.getLastChild();
                                } else {
                                    System.out.println("异常类型");
                                }
                            }
                            if (flag == true) {
                                break;
                            }
                        }

                        if (flag == true) {
                            builder.getCurrentSection().getBody().insertAfter(endTag, table);
                        } else {
                            try {
                                builder.moveTo(currentNode);
                            } catch (Exception e) {
                                builder.moveTo(rangeEnd);
                                builder.insertParagraph();
                                builder.moveTo(rangeEnd);
                            }
                            builder.getCurrentSection().getBody().insertAfter(endTag, builder.getCurrentParagraph());
                        }
                    }

                    rangeStart.remove();
                    rangeEnd.remove();
                } else {
                    builder.moveTo(rangeStart);
                    Node curNode = rangeStart.getNextSibling();
                    ArrayList<Node> nodeList = new ArrayList<>();
                    while (curNode.getNodeType() != NodeType.EDITABLE_RANGE_END) {
                        nodeList.add(curNode);
                        curNode = curNode.getNextSibling();
                    }

                    StructuredDocumentTag sdtRichText = new StructuredDocumentTag(doc, SdtType.RICH_TEXT, MarkupLevel.INLINE);
                    sdtRichText.isShowingPlaceholderText(false);
                    for (int j = 0; j < nodeList.size(); j++) {
                        sdtRichText.getChildNodes().add(nodeList.get(j));
                    }
                    sdtRichText.getRange().replace("Click here to enter text.", "");
                    builder.insertNode(sdtRichText);

                    rangeStart.remove();
                    rangeEnd.remove();
                }
            }

            deleteTitleSpace(doc);

            deleteEmptyPara(doc);

            addMsrList(doc, builder);

            handleSign(doc, builder);

            addContractNo(doc);

            replaceTag(doc, builder);

//            addBuyer(doc, builder);

            changeMsrName(doc);

            addBookmark(doc);

            solveDateHideProblem(doc);

            doc.save(os, SaveFormat.DOCX);
            os.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    static void deleteTitleSpace(Document doc) {
        Section section = doc.getFirstSection();
        ParagraphCollection collection = section.getBody().getParagraphs();
        for (int i = 0; i < collection.getCount(); i++) {
            Paragraph paragraph = collection.get(i);
            if (paragraph.getText().contains("商品房")) {
                // 两个相连的空格
                Run twoSpace = (Run) paragraph.getFirstChild();
                Font font = twoSpace.getFont();
                paragraph.getFirstChild().remove();
                Run space = new Run(doc, " ");
                space.getFont().setSize(font.getSize());
                paragraph.prependChild(space);
                break;
            } else {
                // 如果没有匹配到含有“商品房”字段的标题，则在扫描一定数量段落后跳出循环
                if (i > 25) {
                    break;
                }
            }
        }
    }

    /**
     * 删除目录页中最后的几个空白段落，不要删除分页符
     *
     * @param doc
     */
    static void deleteEmptyPara(Document doc) {
        boolean flag = false;
        Section section = doc.getFirstSection();
        ParagraphCollection collection = section.getBody().getParagraphs();
        for (int i = 0; i < collection.getCount(); i++) {
            Paragraph paragraph = collection.get(i);
            if (paragraph.getText().contains("章")) {
                flag = true;
            }
            if (flag == true) {
                if (paragraph.getText().equals("\r")) {
                    paragraph.remove();
                    i--;
                }
                if (paragraph.getText().contains("\f")) {
                    break;
                }
            }
        }
    }

    static void addMsrList(Document doc, DocumentBuilder builder) {
        boolean flag = false;
        for (Section section : doc.getSections()) {
            ParagraphCollection collection = section.getBody().getParagraphs();
            for (int i = 0; i < collection.getCount(); i++) {
                Paragraph paragraph = collection.get(i);
                if (paragraph.getText().contains("ms_name")) {
                    if (flag == false) {
                        flag = true;
                        continue;
                    } else {
                        Run run = new Run(doc, "{{?msrlist}}");
                        run.getFont().setSize(12);
                        Paragraph startPara = new Paragraph(doc);
                        startPara.appendChild(run);
                        builder.moveTo(paragraph);
                        builder.getCurrentSection().getBody().insertBefore(startPara, paragraph);
                        i++;
                    }
                }
                if (paragraph.getText().contains("ms_agent_tel")) {
                    Run run = new Run(doc, "{{/msrlist}}");
                    run.getFont().setSize(12);
                    Paragraph endPara = new Paragraph(doc);
                    endPara.appendChild(run);
                    builder.moveTo(paragraph);
                    builder.getCurrentSection().getBody().insertAfter(endPara, paragraph);
                    // 由于文档中只有一个{{ms_agent_tel}}标记，所以找到执行后可以直接结束方法
                    return;
                }
            }
        }
    }

    static void handleSign(Document doc, DocumentBuilder builder) throws Exception {
        for (Section section : doc.getSections()) {
            ParagraphCollection collection = section.getBody().getParagraphs();
            for (int i = 0; i < collection.getCount(); i++) {
                Paragraph paragraph = collection.get(i);
                // 删除“网签日期：{{xddate}}”段落
                if (paragraph.getText().contains("xddate")) {
                    Paragraph currentPara = (Paragraph) paragraph.getPreviousSibling();
                    currentPara.getParagraphFormat().setLeftIndent(70);
                    Run run = new Run(doc);
                    run.setText("_R{faren}                 ");
                    run.getFont().setSize(12);
                    run.getFont().setColor(Color.white);
                    currentPara.prependChild(run);

                    while (!currentPara.getText().contains("出卖人")) {
                        currentPara = (Paragraph) currentPara.getPreviousSibling();
                    }
                    currentPara = (Paragraph) currentPara.getPreviousSibling();
                    builder.moveTo(currentPara);
                    Table table = builder.startTable();
                    builder.insertCell();
                    builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
                    builder.getCellFormat().getBorders().setColor(Color.white);
                    builder.getFont().setColor(Color.white);
                    builder.write("_R{seller}");
                    builder.insertCell();
                    builder.getParagraphFormat().setAlignment(ParagraphAlignment.RIGHT);
                    builder.write("_R{buyer}");
                    builder.endRow();
                    builder.endTable();
                    currentPara.remove();

                    break;
                }
            }
        }
    }

    // 在第一页的右上角添加合同编号
    static void addContractNo(Document doc) {
        Paragraph firstPara = doc.getFirstSection().getBody().getFirstParagraph();
        Paragraph secondPara = (Paragraph) firstPara.getNextSibling();
        firstPara.remove();
        secondPara.remove();
        firstPara = doc.getFirstSection().getBody().getFirstParagraph();
        Run run = new Run(doc, "合同编号：{{contract_no}}");
        run.getFont().setSize(14);
        run.getFont().setName("宋体");
        run.getFont().setBold(true);
        Paragraph insertPara = new Paragraph(doc);
        insertPara.getParagraphFormat().setAlignment(ParagraphAlignment.RIGHT);
        insertPara.appendChild(run);
        doc.getFirstSection().getBody().insertBefore(insertPara, firstPara);
    }

    // 将第七条中的“(贷款机构)申请贷款支付”文字前的内容控件替换为双括号标记
    static void replaceTag(Document doc, DocumentBuilder builder) {
        NodeCollection tagNodeCollection = doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true);
        for (int i = 0; i < tagNodeCollection.getCount(); i++) {
            StructuredDocumentTag tag = (StructuredDocumentTag) tagNodeCollection.get(i);
            if ((tag.getNextSibling() != null) && (tag.getNextSibling().getText().contains("(贷款机构)申请贷款支付"))) {
                Node node = tag.getNextSibling();
                tag.remove();
                builder.moveTo(node);
                Run run = new Run(doc, " {{jj_dkfsyhmc}} ");
                run.getFont().setSize(12);
                run.getFont().setName("宋体");
                run.getFont().setUnderline(Underline.SINGLE);
                // 注意insertNode方法是在当前Node的前面插入新Node
                builder.insertNode(run);
                return;
            }
        }
    }

    // 在四个“买受人签字：”文字后面加上_R{buyer}标记
    static void addBuyer(Document doc, DocumentBuilder builder) {
        for (Section section : doc.getSections()) {
            ParagraphCollection collection = section.getBody().getParagraphs();
            for (int i = 0; i < collection.getCount(); i++) {
                Paragraph paragraph = collection.get(i);
                if (paragraph.getText().contains("买受人签字：")) {
                    NodeCollection nodeCollection = paragraph.getChildNodes();
                    for (int j = 0; j < nodeCollection.getCount(); j++) {
                        Node node = nodeCollection.get(j);
                        if ((node != null) && (node.getText().contains("买受人签字："))) {
                            int a;
                            builder.moveTo(node.getNextSibling());
                            Run run = new Run(doc, "_R{buyer}");
                            run.getFont().setSize(12);
                            run.getFont().setName("宋体");
                            run.getFont().setColor(Color.white);
                            builder.insertNode(run);
                        }
                    }
                }
            }
        }
    }

    /**
     * 文档中有两个{{ms_name}}，将第一个{{ms_name}}替换为{{ms_names}}
     */
    static void changeMsrName(Document doc) {
        try {
            FindReplaceOptions options = new FindReplaceOptions();
            options.setReplacingCallback(new ReplaceFirstCallback());
            doc.getRange().replace("{{ms_name}}", "{{ms_names}}", options);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    static class ReplaceFirstCallback implements IReplacingCallback {
        private boolean firstMatch = true;
        // 只有第一次匹配到进行替换，之后匹配到则不替换
        public int replacing(ReplacingArgs e) {
            if (firstMatch) {
                e.setReplacement("{{ms_names}}");
                firstMatch = false;
                return ReplaceAction.REPLACE;
            }
            return ReplaceAction.SKIP;
        }
    }

    /**
     * 在文档中找到“网签日期：{{xddate}}”将{{xddate}}删除，在删除位置添加一个名为“dtag_xydate”的书签
     */
    static void addBookmark(Document doc) {
        try {
            for (Field field : doc.getRange().getFields()) {
                // 查找包含"xddate"字段的域
                if (field.getFieldCode().contains("xddate")) {
                    Node startNode = field.getStart();
                    Node endNode = field.getEnd();

                    // 生成书签
                    DocumentBuilder builder = new DocumentBuilder(doc);
                    BookmarkStart bookmarkStart = builder.startBookmark("dtag_xydate");
                    BookmarkEnd bookmarkEnd = builder.endBookmark("dtag_xydate");
                    // 在域的位置插入书签
                    startNode.getParentNode().insertBefore(bookmarkStart, startNode);
                    endNode.getParentNode().insertAfter(bookmarkEnd, endNode);

                    // 删除域，这里域和“{{xddate}}”都被删除了
                    field.remove();
                    break;
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 第一页最后的日期会被上面的框框遮挡住，需要将标题上面的段落删除几行，日期段落向下移动几行，并在日期后面加上分页符
     */
    static void solveDateHideProblem(Document doc) {
        for (Section section : doc.getSections()) {
            ParagraphCollection collection = section.getBody().getParagraphs();
            int deleteParaNum = 0;
            for (int i = 0; i < collection.getCount(); i++) {
                Paragraph paragraph = collection.get(i);
                // 删除标题之前的五个空段落
                if (paragraph.getText().trim().length() == 0 && deleteParaNum < 5) {
                    paragraph.remove();
                    deleteParaNum++;
                }

                // 找到日期段落
                if (paragraph.getText().contains("月")) {
                    // 在日期段落前面插入两个空段落
                    Paragraph newParagraph1 = new Paragraph(doc);
                    Paragraph newParagraph2 = new Paragraph(doc);
                    collection.insert(i, newParagraph1);
                    collection.insert(i, newParagraph2);

                    // 在日期段落的下一段落添加分页符
                    DocumentBuilder builder = new DocumentBuilder(doc);
                    builder.moveTo(paragraph.getNextSibling());
                    builder.insertBreak(BreakType.PAGE_BREAK);
                    return;
                }
            }
        }
    }
}
