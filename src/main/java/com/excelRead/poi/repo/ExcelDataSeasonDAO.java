package com.excelRead.poi.repo;

import com.excelRead.poi.Entity.CollectionOfSeasonData;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository(value = "mongoExcelSeasonDataDAO")
public interface ExcelDataSeasonDAO extends JpaRepository<CollectionOfSeasonData, String> {

//    List<ExcelEpisodeData> findByProgramType(String s);
}
