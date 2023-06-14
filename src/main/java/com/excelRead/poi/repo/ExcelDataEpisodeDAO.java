package com.excelRead.poi.repo;

import com.excelRead.poi.Entity.CollectionOfEpisodeData;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface ExcelDataEpisodeDAO extends JpaRepository<CollectionOfEpisodeData, String> {

//    List<ExcelEpisodeData> findByProgramType(String s);
}
