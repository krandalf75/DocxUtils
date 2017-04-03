/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package org.krandalf.docx4j.docxutils;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import org.docx4j.openpackaging.contenttype.ContentTypes;

public class FileContentProvider implements ContentProvider {

    
    @Override
    public Content getContent(String target)  {
        try {
            Path path = Paths.get(target);
            byte[] data = Files.readAllBytes(path);
            
            String contentType = "";
            if (target.toLowerCase().endsWith(".jpg")) {
                contentType = ContentTypes.IMAGE_JPEG;
            } else if (target.toLowerCase().endsWith(".jpeg")) {
                contentType = ContentTypes.IMAGE_JPEG;
            } else if (target.toLowerCase().endsWith(".png")) {
                contentType = ContentTypes.IMAGE_PNG;
            } else if (target.toLowerCase().endsWith(".gif")) {
                contentType = ContentTypes.IMAGE_GIF;
            } else if (target.toLowerCase().endsWith(".docx")) {
                contentType = ContentTypes.WORDPROCESSINGML_DOCUMENT;
                //  "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            }
                
           
            return new Content(data,contentType);
        } catch (IOException ex) {
            return null;          
        }
    }
    
}
