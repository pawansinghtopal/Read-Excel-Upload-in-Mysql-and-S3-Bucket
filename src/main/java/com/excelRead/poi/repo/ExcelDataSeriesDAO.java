package com.excelRead.poi.repo;

import com.excelRead.poi.Entity.CollectionOfSeriesData;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository(value = "mongoExcelDataDAO")
public interface ExcelDataSeriesDAO extends JpaRepository<CollectionOfSeriesData, String> {

//    List<ExcelEpisodeData> findByProgramType(String s);
}
