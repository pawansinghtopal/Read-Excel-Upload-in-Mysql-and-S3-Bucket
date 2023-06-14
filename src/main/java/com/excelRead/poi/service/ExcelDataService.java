package com.excelRead.poi.service;

import com.excelRead.poi.Entity.CollectionOfEpisodeData;
import com.excelRead.poi.Entity.CollectionOfSeasonData;
import com.excelRead.poi.Entity.CollectionOfSeriesData;
import com.excelRead.poi.dao.ExcelEpisodeData;
import com.excelRead.poi.dao.ExcelSeasonData;
import com.excelRead.poi.dao.ExcelSeriesData;
import com.excelRead.poi.repo.ExcelDataEpisodeDAO;
import com.excelRead.poi.repo.ExcelDataSeasonDAO;
import com.excelRead.poi.repo.ExcelDataSeriesDAO;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.IOException;
import java.util.Date;
import java.util.Iterator;

@Service
public class ExcelDataService {
    int count = 0;
    @Autowired
    private ExcelDataEpisodeDAO excelDataEpisodeDAO;

    @Autowired
    private ExcelDataSeasonDAO excelDataSeasonDAO;

    @Autowired
    private ExcelDataSeriesDAO excelDataSeriesDAO;


    public void processExcelFiles(String seriesFile, String seasonsFile, String episodeFile) throws IOException {
        processFile(seriesFile);
        processFile(seasonsFile);
        processFile(episodeFile);
    }

    private void processFile(String filepath) throws IOException {
        Workbook workbook;
        File file = new File(filepath);
        String fileName = file.getName();
        String extension = FilenameUtils.getExtension(fileName);
            try {
            if ("xlsx".equals(extension)) {
                System.out.println("xlsx file");
                workbook = new XSSFWorkbook(file);
            } else if ("xls".equals(extension)) {
                System.out.println("xls file");
                workbook = new HSSFWorkbook(POIFSFileSystem.create(file));
            } else {
                throw new RuntimeException();
            }
        }catch (Exception ex) {
            throw new RuntimeException(ex);
        }
        Sheet sheet = workbook.getSheet("Metadata - Web SeriesTV Shows");
        Iterator<Row> rowIterator = sheet.iterator();

        if (rowIterator.hasNext()) {
            rowIterator.next();
        }

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            boolean isRowEmpty = true;
            for (Cell cell : row) {
                if (cell.getCellType() != CellType.BLANK) {
                    isRowEmpty = false;
                    break;
                }
            }
            if (isRowEmpty) {
                return;
            }

            if(fileName.equals("Series.xlsx")){
                ExcelSeriesData excelSeriesData = createExcelSeriesDataFromRow(row);
                CollectionOfSeriesData dataSeries = convertSeriesToEntity(excelSeriesData);
                excelDataSeriesDAO.save(dataSeries);
            }else if(fileName.equals("Season.xlsx")){
                ExcelSeasonData excelSeasonData = createExcelSeasonDataFromRow(row);
                CollectionOfSeasonData dataSeason = convertSeasonToEntity(excelSeasonData);
                excelDataSeasonDAO.save(dataSeason);
            } else if (fileName.equals("Episode.xlsx")) {
                ExcelEpisodeData excelEpisodeData = createExcelEpisodeDataFromRow(row);
                CollectionOfEpisodeData dataEpisode = convertEpisodeToEntity(excelEpisodeData);
                excelDataEpisodeDAO.save(dataEpisode);
            }


        }

        workbook.close();
    }



    private CollectionOfSeriesData convertSeriesToEntity(ExcelSeriesData excelData){

        CollectionOfSeriesData data = new CollectionOfSeriesData();
        BeanUtils.copyProperties(excelData, data);
        return data;
    }
    private CollectionOfSeasonData convertSeasonToEntity(ExcelSeasonData excelData){

        CollectionOfSeasonData data = new CollectionOfSeasonData();
        BeanUtils.copyProperties(excelData, data);
        return data;
    }
    private CollectionOfEpisodeData convertEpisodeToEntity(ExcelEpisodeData excelData) {
        CollectionOfEpisodeData data = new CollectionOfEpisodeData();
        BeanUtils.copyProperties(excelData, data);
        return data;
    }

    private ExcelSeriesData createExcelSeriesDataFromRow(Row row) {
        ExcelSeriesData excelSeriesData = new ExcelSeriesData();
        excelSeriesData.setField(getStringCellValue(row.getCell(0)));
        excelSeriesData.setField(getStringCellValue(row.getCell(1)));
        excelSeriesData.setPartnerName(getStringCellValue(row.getCell(2)));
        excelSeriesData.setUniqueId(getLongCellValue(row.getCell(3)));
        excelSeriesData.setTitle(getStringCellValue(row.getCell(4)));
        excelSeriesData.setProgramType(getStringCellValue(row.getCell(5)));
        excelSeriesData.setDuration(getStringCellValue(row.getCell(6)));
        excelSeriesData.setShortSynopsis(getStringCellValue(row.getCell(7)));
        excelSeriesData.setActors(getStringCellValue(row.getCell(8)));
        excelSeriesData.setDirectors(getStringCellValue(row.getCell(9)));
        excelSeriesData.setReleaseDate(getDateCellValue(row.getCell(10)));
        excelSeriesData.setOriginalLanguage(getStringCellValue(row.getCell(11)));
        excelSeriesData.setOrigin(getStringCellValue(row.getCell(12)));
        excelSeriesData.setGenre(getStringCellValue(row.getCell(13)));
        excelSeriesData.setSeasonNumber(getLongCellValue(row.getCell(14)));
        excelSeriesData.setEpisodeNumber(getLongCellValue(row.getCell(15)));
        excelSeriesData.setRating(getStringCellValue(row.getCell(16)));
        excelSeriesData.setRatingsDescriptor(getStringCellValue(row.getCell(17)));
        excelSeriesData.setTerritory(getStringCellValue(row.getCell(18)));
        excelSeriesData.setStartDate(getDateCellValue(row.getCell(19)));
        excelSeriesData.setEndDate(getDateCellValue(row.getCell(20)));
        excelSeriesData.setBoxartFilename(getStringCellValue(row.getCell(21)));
        excelSeriesData.setCoverartFilename(getStringCellValue(row.getCell(22)));
        excelSeriesData.setHeroartFilename(getStringCellValue(row.getCell(23)));
        excelSeriesData.setAudioLanguage1Code(getStringCellValue(row.getCell(24)));
        excelSeriesData.setAudioLanguage1FileName(getStringCellValue(row.getCell(25)));
        excelSeriesData.setAudioLanguage2Code(getStringCellValue(row.getCell(26)));
        excelSeriesData.setAudioLanguage2FileName(getStringCellValue(row.getCell(27)));
        excelSeriesData.setAudioLanguage3Code(getStringCellValue(row.getCell(28)));
        excelSeriesData.setAudioLanguage3FileName(getStringCellValue(row.getCell(29)));
        excelSeriesData.setAudioLanguage4Code(getStringCellValue(row.getCell(30)));
        excelSeriesData.setAudioLanguage4FileName(getStringCellValue(row.getCell(31)));
        excelSeriesData.setAudioLanguage5Code(getStringCellValue(row.getCell(32)));
        excelSeriesData.setAudioLanguage5FileName(getStringCellValue(row.getCell(33)));
        excelSeriesData.setRunnContentThemes(getStringCellValue(row.getCell(34)));
        excelSeriesData.setRunnContentCategories(getStringCellValue(row.getCell(35)));
        excelSeriesData.setRunnGenreCategories(getStringCellValue(row.getCell(36)));
        excelSeriesData.setContentGroup(getStringCellValue(row.getCell(37)));
        excelSeriesData.setPromotion(getStringCellValue(row.getCell(38)));


        System.out.println("count = "+count);
        count++;
        return excelSeriesData;
    }

    private ExcelSeasonData createExcelSeasonDataFromRow(Row row) {
        ExcelSeasonData excelSeasonData = new ExcelSeasonData();
        excelSeasonData.setField(getStringCellValue(row.getCell(0)));
        excelSeasonData.setField(getStringCellValue(row.getCell(1)));
        excelSeasonData.setPartnerName(getStringCellValue(row.getCell(2)));
        excelSeasonData.setUniqueId(getLongCellValue(row.getCell(3)));
        excelSeasonData.setTitle(getStringCellValue(row.getCell(4)));
        excelSeasonData.setProgramType(getStringCellValue(row.getCell(5)));
        excelSeasonData.setDuration(getStringCellValue(row.getCell(6)));
        excelSeasonData.setShortSynopsis(getStringCellValue(row.getCell(7)));
        excelSeasonData.setActors(getStringCellValue(row.getCell(8)));
        excelSeasonData.setDirectors(getStringCellValue(row.getCell(9)));
        excelSeasonData.setReleaseDate(getDateCellValue(row.getCell(10)));
        excelSeasonData.setOriginalLanguage(getStringCellValue(row.getCell(11)));
        excelSeasonData.setOrigin(getStringCellValue(row.getCell(12)));
        excelSeasonData.setGenre(getStringCellValue(row.getCell(13)));
        excelSeasonData.setSeasonNumber(getLongCellValue(row.getCell(14)));
        excelSeasonData.setEpisodeNumber(getLongCellValue(row.getCell(15)));
        excelSeasonData.setRating(getStringCellValue(row.getCell(16)));
        excelSeasonData.setRatingsDescriptor(getStringCellValue(row.getCell(17)));
        excelSeasonData.setTerritory(getStringCellValue(row.getCell(18)));
        excelSeasonData.setStartDate(getDateCellValue(row.getCell(19)));
        excelSeasonData.setEndDate(getDateCellValue(row.getCell(20)));
        excelSeasonData.setBoxartFilename(getStringCellValue(row.getCell(21)));
        excelSeasonData.setCoverartFilename(getStringCellValue(row.getCell(22)));
        excelSeasonData.setHeroartFilename(getStringCellValue(row.getCell(23)));
        excelSeasonData.setAudioLanguage1Code(getStringCellValue(row.getCell(24)));
        excelSeasonData.setAudioLanguage1FileName(getStringCellValue(row.getCell(25)));
        excelSeasonData.setAudioLanguage2Code(getStringCellValue(row.getCell(26)));
        excelSeasonData.setAudioLanguage2FileName(getStringCellValue(row.getCell(27)));
        excelSeasonData.setAudioLanguage3Code(getStringCellValue(row.getCell(28)));
        excelSeasonData.setAudioLanguage3FileName(getStringCellValue(row.getCell(29)));
        excelSeasonData.setAudioLanguage4Code(getStringCellValue(row.getCell(30)));
        excelSeasonData.setAudioLanguage4FileName(getStringCellValue(row.getCell(31)));
        excelSeasonData.setAudioLanguage5Code(getStringCellValue(row.getCell(32)));
        excelSeasonData.setAudioLanguage5FileName(getStringCellValue(row.getCell(33)));
        excelSeasonData.setRunnContentThemes(getStringCellValue(row.getCell(34)));
        excelSeasonData.setRunnContentCategories(getStringCellValue(row.getCell(35)));
        excelSeasonData.setRunnGenreCategories(getStringCellValue(row.getCell(36)));
        excelSeasonData.setContentGroup(getStringCellValue(row.getCell(37)));
        excelSeasonData.setPromotion(getStringCellValue(row.getCell(38)));
        return excelSeasonData;
    }

    private ExcelEpisodeData createExcelEpisodeDataFromRow(Row row) {
        ExcelEpisodeData excelEpisodeData = new ExcelEpisodeData();
        excelEpisodeData.setField(getStringCellValue(row.getCell(0)));
        excelEpisodeData.setField(getStringCellValue(row.getCell(1)));
        excelEpisodeData.setPartnerName(getStringCellValue(row.getCell(2)));
        excelEpisodeData.setUniqueId(getLongCellValue(row.getCell(3)));
        excelEpisodeData.setTitle(getStringCellValue(row.getCell(4)));
        excelEpisodeData.setProgramType(getStringCellValue(row.getCell(5)));
        excelEpisodeData.setDuration(getStringCellValue(row.getCell(6)));
        excelEpisodeData.setShortSynopsis(getStringCellValue(row.getCell(7)));
        excelEpisodeData.setActors(getStringCellValue(row.getCell(8)));
        excelEpisodeData.setDirectors(getStringCellValue(row.getCell(9)));
        excelEpisodeData.setReleaseDate(getDateCellValue(row.getCell(10)));
        excelEpisodeData.setOriginalLanguage(getStringCellValue(row.getCell(11)));
        excelEpisodeData.setOrigin(getStringCellValue(row.getCell(12)));
        excelEpisodeData.setGenre(getStringCellValue(row.getCell(13)));
        excelEpisodeData.setSeasonNumber(getLongCellValue(row.getCell(14)));
        excelEpisodeData.setEpisodeNumber(getLongCellValue(row.getCell(15)));
        excelEpisodeData.setRating(getStringCellValue(row.getCell(16)));
        excelEpisodeData.setRatingsDescriptor(getStringCellValue(row.getCell(17)));
        excelEpisodeData.setTerritory(getStringCellValue(row.getCell(18)));
        excelEpisodeData.setStartDate(getDateCellValue(row.getCell(19)));
        excelEpisodeData.setEndDate(getDateCellValue(row.getCell(20)));
        excelEpisodeData.setBoxartFilename(getStringCellValue(row.getCell(21)));
        excelEpisodeData.setCoverartFilename(getStringCellValue(row.getCell(22)));
        excelEpisodeData.setHeroartFilename(getStringCellValue(row.getCell(23)));
        excelEpisodeData.setAudioLanguage1Code(getStringCellValue(row.getCell(24)));
        excelEpisodeData.setAudioLanguage1FileName(getStringCellValue(row.getCell(25)));
        excelEpisodeData.setAudioLanguage2Code(getStringCellValue(row.getCell(26)));
        excelEpisodeData.setAudioLanguage2FileName(getStringCellValue(row.getCell(27)));
        excelEpisodeData.setAudioLanguage3Code(getStringCellValue(row.getCell(28)));
        excelEpisodeData.setAudioLanguage3FileName(getStringCellValue(row.getCell(29)));
        excelEpisodeData.setAudioLanguage4Code(getStringCellValue(row.getCell(30)));
        excelEpisodeData.setAudioLanguage4FileName(getStringCellValue(row.getCell(31)));
        excelEpisodeData.setAudioLanguage5Code(getStringCellValue(row.getCell(32)));
        excelEpisodeData.setAudioLanguage5FileName(getStringCellValue(row.getCell(33)));
        excelEpisodeData.setRunnContentThemes(getStringCellValue(row.getCell(34)));
        excelEpisodeData.setRunnContentCategories(getStringCellValue(row.getCell(35)));
        excelEpisodeData.setRunnGenreCategories(getStringCellValue(row.getCell(36)));
        excelEpisodeData.setContentGroup(getStringCellValue(row.getCell(37)));
        excelEpisodeData.setPromotion(getStringCellValue(row.getCell(38)));
        return excelEpisodeData;
    }

    private Date getDateCellValue(Cell cell) {
        if (cell != null){
            cell.setCellValue(CellValue.TRUE.getNumberValue());
            return cell.getDateCellValue();
        }
        return null;
    }

    private String getStringCellValue(Cell cell) {
        if (cell != null) {
            cell.setCellType(CellType.STRING);
            String stringCellValue = cell.getStringCellValue();
            return stringCellValue.trim();
        }
        return null;
    }

    private Long getLongCellValue(Cell cell) {
        if (cell != null) {
            cell.setCellType(CellType.NUMERIC);
            long numericCellValue = (long) cell.getNumericCellValue();
            return numericCellValue;
        }
        return null;
    }

}
