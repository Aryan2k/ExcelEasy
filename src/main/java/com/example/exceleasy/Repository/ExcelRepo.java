package com.example.exceleasy.Repository;

import com.example.exceleasy.Model.Model;
import org.springframework.data.mongodb.repository.MongoRepository;

public interface ExcelRepo extends MongoRepository<Model,String> {
}
