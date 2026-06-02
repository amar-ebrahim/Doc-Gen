package com.bawdocgen.docx;

import com.bawdocgen.api.DocumentGenerationException;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.InputSource;

import javax.xml.XMLConstants;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.ByteArrayOutputStream;
import java.io.StringReader;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashSet;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

class XmlPartTransformer {
    private static final String WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static final String W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";
    private static final Pattern PLACEHOLDER = Pattern.compile("\\{\\{\\s*([A-Za-z0-9_.\\[\\]-]+)\\s*}}");
    private static final Pattern TABLE_PLACEHOLDER = Pattern.compile("\\{\\{\\s*([A-Za-z0-9_.-]+)\\[\\]\\.([A-Za-z0-9_.-]+)\\s*}}");
    private static final Set<String> TRUTHY = new HashSet<>(Arrays.asList("true", "1", "yes", "checked", "on"));

    // String-based split-placeholder patterns (avoids DOM re-serialization which reorders XML attributes).
    // 3-way: Word spell-checker splits {{KEY}} into three <w:t> nodes.
    // Allows leading text (e.g. spaces) before {{ in the first t, and trailing text after }} in the last t.
    // Tried before 2-way so the middle key-t is not consumed by the 2-way lazy match.
    private static final Pattern SPLIT_3WAY = Pattern.compile(
            "(<w:t[^>]*>)([^{<]*)\\{\\{(</w:t>)"               // g1=t1-open, g2=pre-{{ text, g3=</w:t>
            + "((?:(?!</w:p>).)*?)"                              // g4=pre-key XML
            + "(<w:t[^>]*>)([A-Za-z0-9_.\\[\\]-]+)(</w:t>)"    // g5=key-t-open, g6=key, g7=</w:t>
            + "((?:(?!</w:p>).)*?)"                              // g8=post-key XML
            + "(<w:t[^>]*>)\\}\\}([^<]*</w:t>)",                // g9=last-t-open, g10=text-after-}}+</w:t>
            Pattern.DOTALL);
    // 2-way: handles all two-<w:t> splits — key may be split across both elements.
    // Allows leading text before {{ and trailing text after }}.
    // e.g. "spaces{{cust" / "omer.name}}", "{{KEY" / "}}", "{{" / "KEY}}"
    private static final Pattern SPLIT_2WAY = Pattern.compile(
            "(<w:t[^>]*>)([^{<]*)\\{\\{([A-Za-z0-9_.\\[\\]-]*)(</w:t>)" // g1=t1-open, g2=pre-{{, g3=key-start, g4=</w:t>
            + "((?:(?!</w:p>).)*?)"                                        // g5=intermediate XML
            + "(<w:t[^>]*>)([A-Za-z0-9_.\\[\\]-]*)\\}\\}([^<]*</w:t>)",  // g6=t2-open, g7=key-end, g8=text-after-}}+</w:t>
            Pattern.DOTALL);

    byte[] transform(byte[] xmlBytes, Map<String, String> values,
                     Map<String, List<Map<String, String>>> tables) throws DocumentGenerationException {
        String xml = new String(xmlBytes, StandardCharsets.UTF_8);
        // String-based fast path: preserves original XML attribute order and namespace declarations.
        // Used when no structural DOM operations are needed (no table row repeating, no checkboxes).
        if (tables.isEmpty() && !xml.contains("w14:checkbox")) {
            xml = replaceSplitPlaceholders(xml, values);
            xml = replaceSimplePlaceholders(xml, values);
            if (!PLACEHOLDER.matcher(xml).find()) {
                return xml.getBytes(StandardCharsets.UTF_8);
            }
            // Unresolvable split pattern remains — fall through to DOM for correctness.
        }
        Document document = parse(xmlBytes);
        repeatTableRows(document, tables);
        replaceCheckboxControls(document, values);
        replaceFlatPlaceholders(document, values);
        return serialize(document);
    }

    // Handles split-run patterns that Word's spell-checker introduces when placeholder keys
    // contain underscores or camelCase that Word flags as misspellings.
    // SPLIT_3WAY is tried first so its middle key-t is not consumed by the 2-way lazy match.
    private String replaceSplitPlaceholders(String xml, Map<String, String> values) {
        xml = applyPattern(xml, SPLIT_3WAY, values, new SplitReplacer() {
            public String replace(Matcher m, String value) {
                // first-t: pre-{{ text + value; key-t: emptied; last-t: text-after-}} preserved
                return m.group(1) + m.group(2) + value + m.group(3)
                        + m.group(4) + m.group(5) + m.group(7)
                        + m.group(8) + m.group(9) + m.group(10);
            }
            public String key(Matcher m) { return m.group(6); }
        });
        xml = applyPattern(xml, SPLIT_2WAY, values, new SplitReplacer() {
            public String replace(Matcher m, String value) {
                // first-t: pre-{{ text + value; last-t: text-after-}} preserved; intermediate XML kept
                return m.group(1) + m.group(2) + value + m.group(4)
                        + m.group(5) + m.group(6) + m.group(8);
            }
            public String key(Matcher m) { return m.group(3) + m.group(7); }
        });
        return xml;
    }

    private interface SplitReplacer {
        String key(Matcher m);
        String replace(Matcher m, String escapedValue);
    }

    private String applyPattern(String xml, Pattern pattern, Map<String, String> values, SplitReplacer replacer) {
        Matcher m = pattern.matcher(xml);
        StringBuffer sb = new StringBuffer();
        while (m.find()) {
            String key = replacer.key(m);
            String raw = values.containsKey(key) ? values.get(key) : "";
            m.appendReplacement(sb, Matcher.quoteReplacement(replacer.replace(m, xmlEscape(raw))));
        }
        m.appendTail(sb);
        return sb.toString();
    }

    private String replaceSimplePlaceholders(String xml, Map<String, String> values) {
        for (Map.Entry<String, String> entry : values.entrySet()) {
            String placeholder = "{{" + entry.getKey() + "}}";
            String value = xmlEscape(entry.getValue() == null ? "" : entry.getValue());
            xml = xml.replace(placeholder, value);
        }
        return xml;
    }

    private static String xmlEscape(String value) {
        return value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\"", "&quot;");
    }

    Set<String> extractPlaceholders(byte[] xmlBytes) throws DocumentGenerationException {
        Document document = parse(xmlBytes);
        String text = combinedText(document.getDocumentElement());
        Set<String> placeholders = new LinkedHashSet<>();
        Matcher matcher = PLACEHOLDER.matcher(text);
        while (matcher.find()) {
            placeholders.add(matcher.group(1));
        }
        return placeholders;
    }

    private void repeatTableRows(Document document, Map<String, List<Map<String, String>>> tables) {
        List<Element> rows = elementsByLocalName(document, "tr");
        for (Element row : rows) {
            String rowText = combinedText(row);
            Matcher matcher = TABLE_PLACEHOLDER.matcher(rowText);
            if (!matcher.find()) {
                continue;
            }

            String tableName = matcher.group(1);
            List<Map<String, String>> tableRows = tables.get(tableName);
            Node parent = row.getParentNode();
            if (parent == null) {
                continue;
            }

            if (tableRows != null) {
                for (Map<String, String> tableRow : tableRows) {
                    Node clone = row.cloneNode(true);
                    replaceTextNodes(clone, rowScopedValues(tableName, tableRow));
                    parent.insertBefore(clone, row);
                }
            }
            parent.removeChild(row);
        }
    }

    private Map<String, String> rowScopedValues(String tableName, Map<String, String> tableRow) {
        java.util.LinkedHashMap<String, String> values = new java.util.LinkedHashMap<>();
        for (Map.Entry<String, String> entry : tableRow.entrySet()) {
            values.put(tableName + "[]." + entry.getKey(), entry.getValue());
        }
        return values;
    }

    private void replaceFlatPlaceholders(Document document, Map<String, String> values) {
        List<Element> paragraphs = elementsByLocalName(document, "p");
        for (Element paragraph : paragraphs) {
            replaceTextNodes(paragraph, values);
        }
    }

    private void replaceTextNodes(Node container, Map<String, String> values) {
        List<Node> textNodes = wordTextNodes(container);
        if (textNodes.isEmpty()) {
            return;
        }

        boolean changed = false;
        for (Node textNode : textNodes) {
            String original = textNode.getTextContent();
            String replaced = replacePlaceholders(original, values);
            if (!original.equals(replaced)) {
                textNode.setTextContent(replaced);
                changed = true;
            }
        }

        String combined = combinedTextFromNodes(textNodes);
        Matcher unresolved = PLACEHOLDER.matcher(combined);
        if (!unresolved.find()) {
            return;
        }

        String replacedCombined = replacePlaceholders(combined, values);
        if (!combined.equals(replacedCombined) || !changed) {
            textNodes.get(0).setTextContent(replacedCombined);
            for (int i = 1; i < textNodes.size(); i++) {
                textNodes.get(i).setTextContent("");
            }
        }
    }

    private String replacePlaceholders(String text, Map<String, String> values) {
        Matcher matcher = PLACEHOLDER.matcher(text);
        StringBuffer output = new StringBuffer();
        while (matcher.find()) {
            String key = matcher.group(1);
            String replacement = values.get(key);
            if (replacement == null) {
                replacement = "";
            }
            matcher.appendReplacement(output, Matcher.quoteReplacement(replacement));
        }
        matcher.appendTail(output);
        return output.toString();
    }

    // Handles Word content-control checkboxes whose <w:tag w:val="{{key}}"/> matches a flat_mapping key.
    // Sets w14:checked val to 1/0 and updates the visible symbol character in sdtContent.
    private void replaceCheckboxControls(Document document, Map<String, String> values) {
        List<Element> sdts = elementsByLocalName(document, "sdt");
        for (Element sdt : sdts) {
            Element sdtPr = firstChildElement(sdt, WORD_NS, "sdtPr");
            if (sdtPr == null) continue;

            Element checkboxEl = firstChildElement(sdtPr, W14_NS, "checkbox");
            if (checkboxEl == null) continue;

            Element tagEl = firstChildElement(sdtPr, WORD_NS, "tag");
            if (tagEl == null) continue;

            String tagVal = tagEl.getAttributeNS(WORD_NS, "val");
            if (tagVal == null || tagVal.isEmpty()) continue;

            String key = tagVal;
            Matcher m = PLACEHOLDER.matcher(tagVal.trim());
            if (m.matches()) {
                key = m.group(1);
            }

            // Always strip {{}} from tag and alias so no {{ markers remain in the output XML
            tagEl.setAttributeNS(WORD_NS, "w:val", key);
            Element aliasEl = firstChildElement(sdtPr, WORD_NS, "alias");
            if (aliasEl != null) {
                String aliasVal = aliasEl.getAttributeNS(WORD_NS, "val");
                if (aliasVal != null && PLACEHOLDER.matcher(aliasVal.trim()).matches()) {
                    aliasEl.setAttributeNS(WORD_NS, "w:val", key);
                }
            }

            String value = values.get(key);
            if (value == null) continue;

            boolean checked = TRUTHY.contains(value.trim().toLowerCase());

            Element checkedEl = firstChildElement(checkboxEl, W14_NS, "checked");
            if (checkedEl != null) {
                checkedEl.setAttributeNS(W14_NS, "w14:val", checked ? "1" : "0");
            }

            Element sdtContent = firstChildElement(sdt, WORD_NS, "sdtContent");
            if (sdtContent == null) continue;
            List<Node> textNodes = wordTextNodes(sdtContent);
            if (!textNodes.isEmpty()) {
                String symbol = checked
                        ? checkboxChar(checkboxEl, "checkedState", "2611")
                        : checkboxChar(checkboxEl, "uncheckedState", "2610");
                textNodes.get(0).setTextContent(symbol);
            }
        }
    }

    private String checkboxChar(Element checkboxEl, String stateLocalName, String fallbackHex) {
        NodeList children = checkboxEl.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            Node child = children.item(i);
            if (child instanceof Element
                    && stateLocalName.equals(child.getLocalName())
                    && W14_NS.equals(child.getNamespaceURI())) {
                String val = ((Element) child).getAttributeNS(W14_NS, "val");
                if (val != null && !val.isEmpty()) {
                    try {
                        return new String(Character.toChars(Integer.parseInt(val, 16)));
                    } catch (NumberFormatException ignored) { }
                }
            }
        }
        return new String(Character.toChars(Integer.parseInt(fallbackHex, 16)));
    }

    private Element firstChildElement(Node parent, String ns, String localName) {
        NodeList children = parent.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            Node child = children.item(i);
            if (child instanceof Element
                    && localName.equals(child.getLocalName())
                    && ns.equals(child.getNamespaceURI())) {
                return (Element) child;
            }
        }
        return null;
    }

    private Document parse(byte[] xmlBytes) throws DocumentGenerationException {
        try {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            factory.setNamespaceAware(true);
            factory.setExpandEntityReferences(false);
            setFeatureIfSupported(factory, XMLConstants.FEATURE_SECURE_PROCESSING, true);
            setFeatureIfSupported(factory, "http://apache.org/xml/features/disallow-doctype-decl", true);
            setFeatureIfSupported(factory, "http://xml.org/sax/features/external-general-entities", false);
            setFeatureIfSupported(factory, "http://xml.org/sax/features/external-parameter-entities", false);

            DocumentBuilder builder = factory.newDocumentBuilder();
            InputSource source = new InputSource(new StringReader(new String(xmlBytes, StandardCharsets.UTF_8)));
            return builder.parse(source);
        } catch (Exception e) {
            throw new DocumentGenerationException("DOC-002", "Failed to parse DOCX XML part", e);
        }
    }

    private void setFeatureIfSupported(DocumentBuilderFactory factory, String feature, boolean enabled)
            throws ParserConfigurationException {
        try {
            factory.setFeature(feature, enabled);
        } catch (ParserConfigurationException e) {
            if (XMLConstants.FEATURE_SECURE_PROCESSING.equals(feature)) {
                throw e;
            }
        }
    }

    private byte[] serialize(Document document) throws DocumentGenerationException {
        try {
            TransformerFactory transformerFactory = TransformerFactory.newInstance();
            transformerFactory.setFeature(XMLConstants.FEATURE_SECURE_PROCESSING, true);
            Transformer transformer = transformerFactory.newTransformer();
            transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
            transformer.setOutputProperty(OutputKeys.OMIT_XML_DECLARATION, "no");
            transformer.setOutputProperty(OutputKeys.STANDALONE, "yes");

            try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
                transformer.transform(new DOMSource(document), new StreamResult(outputStream));
                return outputStream.toByteArray();
            }
        } catch (Exception e) {
            throw new DocumentGenerationException("DOC-002", "Failed to serialize DOCX XML part", e);
        }
    }

    private List<Element> elementsByLocalName(Document document, String localName) {
        NodeList nodeList = document.getElementsByTagNameNS(WORD_NS, localName);
        List<Element> elements = new ArrayList<>(nodeList.getLength());
        for (int i = 0; i < nodeList.getLength(); i++) {
            Node node = nodeList.item(i);
            if (node instanceof Element) {
                elements.add((Element) node);
            }
        }
        return elements;
    }

    private List<Node> wordTextNodes(Node container) {
        List<Node> textNodes = new ArrayList<>();
        collectWordTextNodes(container, textNodes);
        return textNodes;
    }

    private void collectWordTextNodes(Node node, List<Node> textNodes) {
        if (node == null) {
            return;
        }
        if (isWordTextNode(node)) {
            textNodes.add(node);
            return;
        }
        NodeList children = node.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            collectWordTextNodes(children.item(i), textNodes);
        }
    }

    private boolean isWordTextNode(Node node) {
        return node instanceof Element
                && "t".equals(node.getLocalName())
                && WORD_NS.equals(node.getNamespaceURI());
    }

    private String combinedText(Node container) {
        return combinedTextFromNodes(wordTextNodes(container));
    }

    private String combinedTextFromNodes(List<Node> textNodes) {
        StringBuilder text = new StringBuilder();
        for (Node textNode : textNodes) {
            text.append(textNode.getTextContent());
        }
        return text.toString();
    }
}
