package com.alecode;

import com.alecode.service.XlsService;
import io.micronaut.configuration.picocli.PicocliRunner;
import jakarta.inject.Inject;
import org.bson.Document;
import picocli.CommandLine;

import java.io.File;
import java.io.IOException;
import java.net.URISyntaxException;
import java.util.List;

@CommandLine.Command(name = "demo", description = "...",
        mixinStandardHelpOptions = true)
public class Application implements Runnable {

    @Inject
    private XlsService xlsService;


    public static void main(String[] args) throws Exception {
        PicocliRunner.run(Application.class, args);
    }

    public void run() {
        try {
            List<File> incomingFiles = xlsService.getFilesInResourcesDirectory();
            List<Document> allDocuments = incomingFiles.stream()
                    .flatMap(f -> xlsService.convertXlsToDocuments(f).stream())
                    .toList();
            xlsService.saveDocuments(allDocuments);
        } catch (IOException ex) {
            throw new RuntimeException(ex);
        } catch (URISyntaxException ex) {
            throw new RuntimeException(ex);
        };
    }
}