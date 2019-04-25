package com.fteretriever.dao;

import java.util.Date;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.data.mongodb.core.MongoOperations;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.stereotype.Repository;

import com.fteretriever.model.FTERecord;

@Repository(value = "FteRetireve")
public abstract class FteRetrieverRepoImpl implements FteRetrieverRepo {
}
