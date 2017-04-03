/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package org.krandalf.docx4j.docxutils;

import com.sampullara.cli.Args;
import com.sampullara.cli.Argument;
import java.io.File;
import java.util.ArrayList;
import java.util.List;
import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.model.fields.ComplexFieldLocator;
import org.docx4j.model.fields.FieldRef;
import org.docx4j.model.fields.FieldsPreprocessor;
import org.docx4j.openpackaging.contenttype.ContentType;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.WordprocessingML.AlternativeFormatInputPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.Body;
import org.docx4j.wml.CTAltChunk;
import org.docx4j.wml.P;
import org.docx4j.wml.Text;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 *
 */
public class DocumentEmbedder {

    private static final Logger log = LoggerFactory.getLogger(DocumentEmbedder.class);

    @Argument(alias = "input", description = "input", required = true)
    private static String inputfilepath;

    @Argument(alias = "output", description = "output", required = true)
    private static String outputfilepath;

    /**
     *
     * @param args
     * @throws java.lang.Exception
     */
    public static void main(String[] args) throws Exception {

        try {
            Args.parse(DocumentEmbedder.class, args);
        } catch (IllegalArgumentException e) {
            Args.usage(DocumentEmbedder.class);
            System.exit(-1);
        }
        File inputFile = new File(inputfilepath);
        System.out.println(inputFile.getAbsolutePath());

        WordprocessingMLPackage word = WordprocessingMLPackage.load(inputFile);
        documentEmbed(word, new FileContentProvider());

        File outputFile = new File(outputfilepath);
        word.save(outputFile);
    }

    public static void documentEmbed(WordprocessingMLPackage word, ContentProvider provider) throws Docx4JException {
        MainDocumentPart documentPart = word.getMainDocumentPart();
        org.docx4j.wml.Document wmlDocumentEl = (org.docx4j.wml.Document) documentPart.getContents();
        Body body = wmlDocumentEl.getBody();

        // find fields
        ComplexFieldLocator fl = new ComplexFieldLocator();
        new TraversalUtil(body, fl);

        int numdoc = 1;
        for (P p : fl.getStarts()) {
            List<FieldRef> fieldRefs = new ArrayList<>();
            P aux = FieldsPreprocessor.canonicalise(p, fieldRefs);

            for (FieldRef fr : fieldRefs) {
                String instr = extractInstr(fr.getInstructions());
                log.info("INSTR:" + instr);
                if (instr.contains("INCLUDETEXT")) {
                    String docfile = getFilenameFromInstr(instr);
                    Content content = provider.getContent(docfile);
                    log.info("docfile:" + docfile);
                    replaceDocx(documentPart, p, content.getData(), numdoc++);
                }
            }
            

        }

    }

    /**
     * Get the datafield name from, for example
     *
     * @param instr
     * @return
     */
    protected static String getFilenameFromInstr(String instr) {

        String tmp = instr.substring(instr.indexOf("INCLUDETEXT") + 11);
        tmp = tmp.trim();
        String datafieldName;
        // A data field name will be quoted if it contains spaces
        if (tmp.startsWith("\"")) {
            if (tmp.indexOf("\"", 1) > -1) {
                datafieldName = tmp.substring(1, tmp.indexOf("\"", 1));
            } else {
                log.warn("Quote mismatch in " + instr);
                // hope for the best
                datafieldName = tmp.contains(" ") ? tmp.substring(1, tmp.indexOf(" ")) : tmp.substring(1);
            }
        } else {
            datafieldName = tmp.contains(" ") ? tmp.substring(0, tmp.indexOf(" ")) : tmp;
        }
        log.info("Key: '" + datafieldName + "'");

        return datafieldName;

    }

    protected static String extractInstr(List<Object> instructions) {
        // For MERGEFIELD, expect the list to contain a simple string

        if (instructions.size() != 1) {
            log.error("TODO MERGEFIELD field contained complex instruction");
            return null;
        }

        Object o = XmlUtils.unwrap(instructions.get(0));
        if (o instanceof Text) {
            return ((Text) o).getValue();
        } else {
            if (log.isErrorEnabled()) {
                log.error("TODO: extract field name from " + o.getClass().getName());
                log.error(XmlUtils.marshaltoString(instructions.get(0), true, true));
            }
            return null;
        }
    }

    private static void replaceDocx(MainDocumentPart main, P p, byte[] bytes, int chunkId) {
        try {
            AlternativeFormatInputPart afiPart = new AlternativeFormatInputPart(new PartName("/part" + chunkId + ".docx"));
            afiPart.setContentType(new ContentType("application/vnd.openxmlformats-officedocument.wordprocessingml.document"));
            afiPart.setBinaryData(bytes);
            Relationship altChunkRel = main.addTargetPart(afiPart);

            CTAltChunk chunk = Context.getWmlObjectFactory().createCTAltChunk();
            chunk.setId(altChunkRel.getId());

            //main.addObject(chunk);
            
            org.docx4j.wml.Document wmlDocumentEl = (org.docx4j.wml.Document) main.getContents();
            Body body = wmlDocumentEl.getBody();
            int position = body.getContent().indexOf(p);
            if (position  >= 0) {
                body.getContent().set(position, chunk);
            } else {
                System.out.println("OOOOOOOOOOOOOOOOPPPPPPPPPPPPPPSSSSSSSSSS");
            }
            
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
