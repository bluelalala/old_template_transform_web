package com.example.old_template_transform_web;

import com.aspose.words.*;
import com.aspose.words.Shape;

import java.awt.*;
import java.io.*;
import java.lang.reflect.Modifier;
import java.util.ArrayList;

public class TransformUtil {

    /**
     * 将域（被替换的值）转换为双括号形式，转换复选框
     */
    public static byte[] procCustomDocument(InputStream stream) {
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
            Document doc = new Document(stream);

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

            removeHeader(doc);

            addHeader(doc);

            doc.save(os, SaveFormat.DOCX);
            bytes = os.toByteArray();
            os.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return bytes;
    }

    static void removeHeader(Document doc) {
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
        // 设置除第一页外的页眉页脚
        builder.getPageSetup().setDifferentFirstPageHeaderFooter(false);
        // 设置奇数页和偶数页页眉页脚
        builder.getPageSetup().setOddAndEvenPagesHeaderFooter(false);
        // 2、开始插入页脚
        // 将光标移动到页脚位置
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        //靠右对齐
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.RIGHT);
        //   设置页脚上下边距
//        builder.getPageSetup().setHeaderDistance(40);
//        builder.getPageSetup().setFooterDistance(0);


        Paragraph paragraph = builder.insertParagraph();

        Run run = new Run(doc, "{{qrCode_ewm}}");
        // 字号小五
//        run.getFont().setSize(9);
//        run.getFont().setName("宋体");

        paragraph.appendChild(run);

        // 添加页眉线
        Border borderHeader = builder.getParagraphFormat().getBorders().getBottom();
        borderHeader.setShadow(true);
        borderHeader.setDistanceFromText(2);
        borderHeader.setLineStyle(LineStyle.SINGLE);
    }

    /**
     * 将可编辑域转换为格式文本内容控件
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
                        builder.insertNode(new Run(doc, "\f"));
                    }

                    if ((rangeEnd.getNextSibling() != null) && (rangeEnd.getNextSibling().getText().equals("\f"))
                            && (rangeEnd.getPreviousSibling() != null)) {
                        builder.moveTo(rangeEnd.getNextSibling());
                        builder.insertParagraph();

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

            addMsrList(doc, builder);

            handleSign(doc, builder);

            addContractNo(doc);

            replaceTag(doc, builder);

//            addBuyer(doc, builder);

            doc.save(os, SaveFormat.DOCX);
            os.close();
        } catch (Exception e) {
            e.printStackTrace();
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

    static void handleSign(Document doc, DocumentBuilder builder) {
        for (Section section : doc.getSections()) {
            ParagraphCollection collection = section.getBody().getParagraphs();
            for (int i = 0; i < collection.getCount(); i++) {
                Paragraph paragraph = collection.get(i);
                // 删除“网签日期：{{xddate}}”段落
                if (paragraph.getText().contains("xddate")) {
                    paragraph.remove();
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
}
