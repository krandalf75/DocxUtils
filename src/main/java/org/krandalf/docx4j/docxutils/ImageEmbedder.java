/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package org.krandalf.docx4j.docxutils;

import com.sampullara.cli.Args;
import com.sampullara.cli.Argument;
import java.io.File;
import java.util.List;
import org.docx4j.dml.CTBlip;
import org.docx4j.openpackaging.contenttype.ContentType;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.relationships.Relationships;
import org.docx4j.utils.SingleTraversalUtilVisitorCallback;
import org.docx4j.utils.TraversalUtilVisitor;
import org.docx4j.wml.Body;

/**
 *
 * @author arovira
 */
public class ImageEmbedder {

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
            Args.parse(ImageEmbedder.class, args);
        } catch (IllegalArgumentException e) {
            Args.usage(ImageEmbedder.class);
            System.exit(-1);
        }
        File inputFile = new File(inputfilepath);

        WordprocessingMLPackage word = WordprocessingMLPackage.load(inputFile);
        imageEmbed(word,new FileContentProvider());

        File outputFile = new File(outputfilepath);
        word.save(outputFile);
    }

    public static void imageEmbed(WordprocessingMLPackage word, ContentProvider provider) throws Docx4JException
    {
        MainDocumentPart documentPart = word.getMainDocumentPart();
        org.docx4j.wml.Document wmlDocumentEl = (org.docx4j.wml.Document) documentPart.getContents();
        Body body = wmlDocumentEl.getBody();

        RelationshipsPart relsPart = documentPart.getRelationshipsPart();
        Relationships rels = relsPart.getRelationships();
        List<Relationship> relsList = rels.getRelationship();
        
        // For each image rel
        int i = 0;
        while (i< relsList.size()) {
            Relationship r = relsList.get(i);
            i++;
            if (r.getType().equals(Namespaces.IMAGE)) {
                System.out.println(r.getId());
                System.out.println(r.getTargetMode());
                if (r.getTargetMode() != null && r.getTargetMode().equals("External")) {
                    relsList.remove(i-1);
                    String target = r.getTarget();
                    String partName = target.startsWith("/")?target:"/"+target;
                    System.out.println("target: " + target);
                    System.out.println("partName: " + partName);
                    BinaryPart imagePart = new BinaryPart(new PartName(partName));
                    Content content = provider.getContent(target);
                    imagePart.setBinaryData(content.getData());
                    imagePart.setContentType(new ContentType(content.getContentType()));
                    imagePart.setRelationshipType(Namespaces.IMAGE);
                    Relationship re = documentPart.addTargetPart(imagePart);
                    re.setId(r.getId());
                    System.out.println(re.getId());
                    i=0;
                }
            }
        }

        SingleTraversalUtilVisitorCallback imageVisitor
                = new SingleTraversalUtilVisitorCallback(
                        new ImageEmbedder.TraversalUtilBlipVisitor());

        imageVisitor.walkJAXBElements(body);
         
    }
    
    
    /**
     * Transform link to embedded
     */
    public static class TraversalUtilBlipVisitor extends TraversalUtilVisitor<CTBlip> {

        @Override
        public void apply(CTBlip element, Object parent, List<Object> siblings) {

            if (element.getLink() != null) {
                System.out.println("LINK:" + element.getLink());
                String relId = element.getLink();
                // Add r:embed
                element.setLink(null);
                // Remove r:link
                element.setEmbed(relId);
            }
        }

    }

}
