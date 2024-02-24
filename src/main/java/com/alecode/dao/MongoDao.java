package com.alecode.dao;

import com.alecode.service.XlsService;
import com.mongodb.client.MongoClient;
import com.mongodb.client.MongoCollection;
import com.mongodb.client.result.InsertManyResult;
import jakarta.inject.Inject;
import org.bson.Document;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;

public class MongoDao {

    private static final Logger LOG = LoggerFactory.getLogger(XlsService.class);

    @Inject
    MongoClient mongoClient;

    public void writeDocuments(List<Document> documents){
        MongoCollection<Document> collection = mongoClient.getDatabase("local").getCollection("xls");
        InsertManyResult insertManyResult = collection.insertMany(documents);
        LOG.info(String.format("Successfully inserted %s records into local:xls", insertManyResult.getInsertedIds().size()));
    }
}
