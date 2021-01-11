package fudala.mateusz.eml;

import com.aspose.email.Attachment;
import com.aspose.email.AttachmentCollection;
import com.aspose.email.EmlLoadOptions;
import com.aspose.email.MailMessage;

import java.io.File;
import java.io.IOException;

public class Main {
    public static void main(String[] args) throws IOException {
        File directory = new File("c:\\messages\\");
        File[] files = directory.listFiles();

        int indexFolder = 0;
        for (File file : files) {
            String strBaseFolder = "C:\\messages\\";
            System.out.println("File name: " + file.getName());
            String fileName = strBaseFolder + file.getName();
            try {
                MailMessage message = MailMessage.load(fileName, new EmlLoadOptions());
                AttachmentCollection attachments = message.getAttachments();
                System.out.println("Attachment Count: " + attachments.size());

                for (int i = 0; i < attachments.size(); i++) {
                    Attachment attachment = attachments.get_Item(i);
                    System.out.println(attachment.getName());
                    attachment.save("c:\\messages\\output\\" + indexFolder + "_" + attachment.getName());
                }
            } catch (Exception e) {
                continue;
            }
            indexFolder++;
        }
    }
}