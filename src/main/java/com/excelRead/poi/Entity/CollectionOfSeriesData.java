package com.excelRead.poi.Entity;

import com.excelRead.poi.dao.ExcelSeasonData;
import jakarta.persistence.*;
import lombok.*;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.UUID;

@Entity
@Table(name = "excel_series_data")
@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
public class CollectionOfSeriesData {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long series_id;
    private String field;
    private String partnerName;
    private Long uniqueId;
    private String title;
    private String programType;
    private String duration;
    private String shortSynopsis;
    private String actors;
    private String directors;
    private Date releaseDate;
    private String originalLanguage;
    private String origin;
    private String genre;
    private Long seasonNumber;
    private Long episodeNumber;
    private String rating;
    private String ratingsDescriptor;
    private String territory;
    private Date startDate;
    private Date endDate;
    private String boxartFilename;
    private String coverartFilename;
    private String heroartFilename;
    private String audioLanguage1Code;
    private String audioLanguage1FileName;
    private String audioLanguage2Code;
    private String audioLanguage2FileName;
    private String audioLanguage3Code;
    private String audioLanguage3FileName;
    private String audioLanguage4Code;
    private String audioLanguage4FileName;
    private String audioLanguage5Code;
    private String audioLanguage5FileName;
    private String runnContentThemes;
    private String runnContentCategories;
    private String runnGenreCategories;
    private String contentGroup;
    private String promotion;

    @OneToMany(mappedBy = "collectionOfSeriesData", cascade = CascadeType.ALL, orphanRemoval = true)
    private List<CollectionOfSeasonData> collectionOfSeasonData = new ArrayList<>();

    public void addSeason(CollectionOfSeasonData season) {
        if(collectionOfSeasonData==null)
        {
            collectionOfSeasonData.add(season);
        }
        season.setCollectionOfSeriesData(this);
    }


}
