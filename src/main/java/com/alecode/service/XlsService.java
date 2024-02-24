package com.alecode.service;


import com.alecode.dao.MongoDao;
import io.micronaut.core.annotation.NonNull;
import jakarta.inject.Singleton;
import jakarta.validation.constraints.NotNull;
import org.apache.poi.ss.usermodel.*;
import org.bson.Document;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

@Singleton
public class XlsService {
    private static final Logger LOG = LoggerFactory.getLogger(XlsService.class);

    private MongoDao mongoDao;

    public XlsService(MongoDao mongoDao) {
        this.mongoDao = mongoDao;
    }

    public void writeXlsToMongo(@NotNull File file ){
        List<Document> documents = convertXlsToDocuments(file);
        mongoDao.writeDocuments(documents);
    }

    public void saveDocuments(List<Document> documents ){
        mongoDao.writeDocuments(documents);
    }

    @NonNull
    public List<Document> convertXlsToDocuments(@NonNull File file ) {
        try {
            LOG.info( "Reading " + file.toString());
            List<Map<String, Object>> data = readExcel(file);
            if (!data.isEmpty()) {
                List<Document> xlsRecords = data.stream()
                        .map(Document::new)
                        .toList();
                LOG.info("Conversion successful. " + xlsRecords.size() + " xls records");
                return xlsRecords;
            } else {
                LOG.warn("No data found in the Excel file. " + file.toString());
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return new ArrayList<>();
    }


    public List<File> getFilesInResourcesDirectory() throws IOException, URISyntaxException {
        Stream<Path> incomingPaths = Files.walk(Paths.get("incoming/").toAbsolutePath());
        return incomingPaths
                .filter( p -> p.toString().endsWith(".xlsx"))
                .map(Path::toFile).toList();
    }


    private static List<Map<String, Object>> readExcel(File file) throws IOException {
        List<Map<String, Object>> data = new ArrayList<>();
        FileInputStream fis = new FileInputStream(file);
        Workbook workbook = WorkbookFactory.create(fis);
        Sheet sheet = workbook.getSheetAt(0);// Assuming you want the first sheet
        Iterator<Row> rowIterator = sheet.iterator();
        if( !rowIterator.hasNext() ){
            LOG.error( "Empty first row on sheet 0 " + file.getName());
            workbook.close();
            fis.close();
            return new ArrayList<>();
        }

        Row rowKeys = rowIterator.next();
        List<String> keys = StreamSupport.stream(rowKeys.spliterator(), false)
                .map(XlsService::getCellStringValue)
                .toList();

        while (rowIterator.hasNext()) {

            Row row = rowIterator.next();
            Map<String, Object> resultRowMap = new HashMap<>();
            for( int i = 0; i < keys.size(); i++ ){
                String tempKey = keys.get(i);
                Cell tempCell = row.getCell(i);
                String tempValue = tempCell != null ? getCellStringValue(tempCell) : "";
                resultRowMap.put(tempKey, tempValue);
            }
            resultRowMap.put("filename", file.getName());
            data.add(resultRowMap);
        }

        workbook.close();
        fis.close();
        return data;
    }

    public static String getCellStringValue(Cell cell){
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            default -> "";
        };
    }
}
