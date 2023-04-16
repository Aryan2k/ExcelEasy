package com.example.exceleasy.Repository;

import com.example.exceleasy.Model.SubModel;
import org.springframework.data.mongodb.repository.MongoRepository;

public interface SubRepo extends MongoRepository<SubModel,String> {
}
