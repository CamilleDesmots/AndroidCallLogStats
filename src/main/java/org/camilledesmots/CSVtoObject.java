/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package org.camilledesmots;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileFilter;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.time.LocalDateTime;
import java.time.ZoneOffset;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.DoubleSummaryStatistics;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.function.BiConsumer;
import java.util.function.Function;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.stream.Collectors;
import org.camilledesmots.androidcalllogstats.CallLogItem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.time.DayOfWeek;

/**
 * A partir de l'historique des sauvegardes Androïd avec l'extension ".clbu"
 * créé des statistiques sous la forme d'un fichier EXCEL.
 *
 * @author camilledesmots
 */
public class CSVtoObject {

    private static final String COMMA = ";";

    private List<CallLogItem> callLogItemList;
    
    private String clbuFolder;
    
    private String resultFolder;

    public String getClbuFolder() {
        return clbuFolder;
    }

    public void setClbuFolder(String clbuFolder) {
        this.clbuFolder = clbuFolder;
    }

    public String getResultFolder() {
        return resultFolder;
    }

    public void setResultFolder(String resultFolder) {
        this.resultFolder = resultFolder;
    }
    

    public CSVtoObject() {
        callLogItemList = new ArrayList<CallLogItem>();
    }

    // Create a Logger 
    private static final Logger LOG = Logger.getLogger(CSVtoObject.class.getName());

    public void readDirectory(String clbuFolder, String resultFolder) throws IOException {
        //Creating a File object for directory
        File directoryPath = new File(clbuFolder);
        LOG.log(Level.INFO, "Lecture du répertoire \"" + clbuFolder + "\"");
        FileFilter textFilefilter;
        textFilefilter = new FileFilter() {
            @Override
            public boolean accept(File file) {

                if (file.getName().endsWith(".clbu")) {
                    return file.isFile();
                } else {
                    return Boolean.FALSE;
                }
            }
        };

        //List of all the text files
        File filesList[] = directoryPath.listFiles(textFilefilter);
        LOG.log(Level.INFO, "Liste des fichiers du répertoire :");
        for (File file : filesList) {
            LOG.log(Level.INFO, "Nom du fichier : {0}", new Object[]{file.getName()});
        }

        for (File file : filesList) {
            this.callLogItemList.addAll(this.processInputFile(file.getAbsolutePath()));
        }

        LOG.log(Level.INFO, "Nombre d'éléments au total :" + this.callLogItemList.size());

        // Date maximum et minimum
        Comparator<CallLogItem> comparator = Comparator.comparing(CallLogItem::getLocalDateTime);

        // Tri de la liste
        Collections.sort(this.callLogItemList, comparator);

        CallLogItem min = this.callLogItemList.stream().min(comparator).get();
        CallLogItem max = this.callLogItemList.stream().max(comparator).get();

        // Set intermédiaire pour supprimer les doubles
        Set<Long> setDate = new HashSet<>();

        this.callLogItemList.removeIf(p -> setDate.add(p.getLocalDateTime().toEpochSecond(ZoneOffset.UTC)) == Boolean.FALSE);

        LOG.log(Level.INFO, "Nombre d'éléments au total sans double : " + this.callLogItemList.size());

        CallLogItem minAfter = this.callLogItemList.stream().min(comparator).get();
        CallLogItem maxAfter = this.callLogItemList.stream().max(comparator).get();

        LOG.log(Level.INFO, "Min : " + min.getLocalDateTime());
        // LOG.log(Level.INFO, "Min après suppression de doublons : " + minAfter.getLocalDateTime());
        LOG.log(Level.INFO, "Max : " + max.getLocalDateTime());
        // LOG.log(Level.INFO, "Max après suppression de doublons : " + maxAfter.getLocalDateTime());

        // Création du fichier au format EXCEL xls
        DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy-MM-dd_HH-mm-ss");
        LocalDateTime now = LocalDateTime.now();

        String fileName = resultFolder + "/Statistiques des appels au "
                + max.getDAYFromLocalDateTime().toString().substring(0, 10)
                + " généré le " + dtf.format(now)
                + ".xlsx";
        LOG.log(Level.INFO, "Création du fichier \"" + fileName + "\"");

        // Format XLSX Apache
        Workbook wb = new XSSFWorkbook();
        this.workSheetSyntheseXlsx(wb);
        this.worksheetDonneesXlsx(wb);
        this.worksheetAppelsParJourXlsx(wb);
        this.worksheetAppelsParMoisXlsx(wb);
        this.worksheetAppelsParAnXlsx(wb);
        this.worksheetAppelsPar24HXlsx(wb);
        this.worksheetAppelsParJourDeLaSemaineXlsx(wb);

        this.worksheetAppelsParJourEntrantXlsx(wb);
        this.worksheetAppelsParMoisEntrantXlsx(wb);
        this.worksheetAppelsParAnEntrantXlsx(wb);
        this.worksheetAppelsPar24HEntrantsXlsx(wb);
        this.worksheetAppelsParJourDeLaSemaineEntrantsXlsx(wb);

        this.worksheetAppelsParJourSortantXlsx(wb);
        this.worksheetAppelsParMoisSortantXlsx(wb);
        this.worksheetAppelsParAnSortantXlsx(wb);
        this.worksheetAppelsPar24HSortantsXlsx(wb);
        this.worksheetAppelsParJourDeLaSemaineSortantsXlsx(wb);

        // Write the output to a file
        try ( FileOutputStream fileOut = new FileOutputStream(fileName)) {
            wb.write(fileOut);
        } catch (IOException ex) {
            LOG.log(Level.SEVERE, null, ex);
        }
    }

    /**
     * A partir d'un fichier retourne une liste d'élément de la Classe
     * CallLogItem
     *
     * @param inputFilePath
     * @return List<CallLogItem>
     */
    private List<CallLogItem> processInputFile(String inputFilePath) {

        LOG.log(Level.INFO, "Nom du fichier : {0}", new Object[]{inputFilePath});
        List<CallLogItem> inputList = new ArrayList<>();

        try {
            File inputF = new File(inputFilePath);
            InputStream inputFS = new FileInputStream(inputF);
            // Skip the header
            BufferedReader br = new BufferedReader(new InputStreamReader(inputFS));
//            br.lines().forEach(e ->{
//                LOG.info("item " + e.toString());
//                LOG.info("Split length " + e.toString().split(COMMA).length);
//            });
            //LOG.log(Level.INFO, "Nombre d'enregistrements à lire dans le fichier " + br.lines().count());
            inputList = br.lines().map(mapToItem).collect(Collectors.toList());
            br.close();

            LOG.log(Level.INFO, "Nombre d'enregistrements transformé en classes depuis le fichier \"" + inputFilePath + "\" : " + inputList.size());
//            LOG.info("--------------------------------------------");
//            inputList.forEach(item -> {
//                LOG.info("CallLogItem " + item.toString());
//            });

        } catch (FileNotFoundException e) {
            // TODO Log Error here
            LOG.log(Level.SEVERE, null, e);

        } catch (IOException e) {
            // TODO Log Error here
            LOG.log(Level.SEVERE, null, e);

        }

        Comparator<CallLogItem> comparator = Comparator.comparing(CallLogItem::getLocalDateTime);

        CallLogItem min = inputList.stream().min(comparator).get();
        CallLogItem max = inputList.stream().max(comparator).get();

        LOG.log(Level.INFO, "   Min " + min.getLocalDateTime());
        LOG.log(Level.INFO, "   Max " + max.getLocalDateTime());

        return inputList;
    }

    /**
     * Function that from a String return a class CallLogItem
     */
    private final Function<String, CallLogItem> mapToItem = (line) -> {

        String[] p = line.split(COMMA);
        //LOG.info(">> Ligne \" + line + ”\" split length : " + p.length);
        CallLogItem item = new CallLogItem();

        item.setNumber(p[0]);
        item.setType(Integer.decode(p[1]));
        item.setLocalDateTime(Long.decode(p[2]));
        item.setDuration(Long.decode(p[3]));

        // LOG.info("Days " + item.getDAYFromLocalDateTime());
        // LOG.info("Month " + item.getMONTHFromLocalDateTime());
        return item;
    };

    /**
     * Crée un onglet nommé "Synthèse à intégrer dans le document XML
     *
     * @return Une description XML d'un ExcelWorksheet EXCEL au format XML.
     */
    private void workSheetSyntheseXlsx(Workbook wb) {

        String ongletName;
        ongletName = "Synthèse";
        LOG.log(Level.INFO, "Création de l'onglet \"" + ongletName + "\"");

        CreationHelper creationHelper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet(ongletName);

        Row row_0 = sheet.createRow(0);
        row_0.createCell(0).setCellValue("Nombre d'enregistremnts traités");
        row_0.createCell(1).setCellValue(this.callLogItemList.size());

        // Date maximum et minimum
        Comparator<CallLogItem> comparator = Comparator.comparing(CallLogItem::getLocalDateTime);
        CallLogItem min = this.callLogItemList.stream().min(comparator).get();
        CallLogItem max = this.callLogItemList.stream().max(comparator).get();

        LOG.log(Level.INFO, "Min " + min.getLocalDateTime());
        LOG.log(Level.INFO, "Max " + max.getLocalDateTime());

        DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss.SSS");

        CellStyle style = wb.createCellStyle();
        // Date au format 2019-05-05T19:26:17.394
        style.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy/mm/dd hh:mm:ss.SSS"));

        Row row_1 = sheet.createRow(1);
        row_1.createCell(0).setCellValue("Date du plus ancien appel");
        Cell cell_1_1 = row_1.createCell(1);
        cell_1_1.setCellValue(dtf.format(min.getLocalDateTime()));
        cell_1_1.setCellStyle(style);

        Row row_2 = sheet.createRow(2);
        row_2.createCell(0).setCellValue("Date du plus récent appel");
        Cell cell_2_1 = row_2.createCell(1);
        cell_2_1.setCellValue(dtf.format(max.getLocalDateTime()));
        cell_2_1.setCellStyle(style);
    }

    /**
     * Détails des données avec numéro de téléphone sans les 2 derniers
     * chiffres.
     *
     * @return Une description XML d'un ExcelWorksheet EXCEL au format XML.
     */
    private void worksheetDonneesXlsx(Workbook wb) {
        String ongletName = "Données";

        short rowCpt = 0;

        LOG.log(Level.INFO, "Création de l'onglet \"" + ongletName + "\"");

        CreationHelper creationHelper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet(ongletName);

        // Titre de colonnes
        //TODO Appliquer un style et une mise en forme à la 1ère ligne 
        Row row_0 = sheet.createRow(rowCpt);

        Cell cell_0_0 = row_0.createCell(0);
        cell_0_0.setCellValue("Date");

        Cell cell_0_1 = row_0.createCell(1);
        cell_0_1.setCellValue("Numéro");

        Cell cell_0_2 = row_0.createCell(2);
        cell_0_2.setCellValue("Type d'appel");

        Cell cell_0_3 = row_0.createCell(3);
        cell_0_3.setCellValue("Durée d'appel en secondes");

        // Transformation de la Map en ExcelExcelRow 
        // Date initialement au format 2019-05-05T19:26:17.394
        DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss.SSS");

        CellStyle style = wb.createCellStyle();
        // Date au format 2019-05-05T19:26:17.394
        style.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy/mm/dd hh:mm:ss.SSS"));

        for (CallLogItem i : this.callLogItemList) {
            rowCpt++;
            Row row_n = sheet.createRow(rowCpt);

            Cell cell_n_0 = row_n.createCell(0);
            cell_n_0.setCellValue(dtf.format(i.getLocalDateTime()));
            cell_n_0.setCellStyle(style);

            row_n.createCell(1).setCellValue(i.getNumber().replaceFirst("\\d{2}$", "XX"));
            row_n.createCell(2).setCellValue(i.getType());
            row_n.createCell(3).setCellValue(i.getDuration());
        }
    }

    /**
     * Statistique des appels par jour.
     *
     * @return Une description XML d'un ExcelWorksheet EXCEL au format XML.
     */
    private void worksheetAppelsParJourXlsx(Workbook wb) {
        String ongletName = "Appels par jour";

        short rowCpt = 0;

        LOG.log(Level.INFO, "Création de l'onglet \"" + ongletName + "\"");

        CreationHelper creationHelper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet(ongletName);

        // Titre de colonnes
        //TODO Appliquer un style et une mise en forme à la 1ère ligne 
        Row row_0 = sheet.createRow(rowCpt);

        row_0.createCell(0).setCellValue("Date");
        row_0.createCell(1).setCellValue("Nombre d'appels entrants");
        row_0.createCell(2).setCellValue("Durée d'appels en minutes");
        row_0.createCell(3).setCellValue("Moyenne temps d'appels en minutes");
        row_0.createCell(4).setCellValue("Min temps d'appels en minutes");
        row_0.createCell(5).setCellValue("Max temps d'appels en minutes");

        // Statistiques par jour
        // Un objet DoubleSummaryStatistics contient 
        // - count
        // - sum
        // - min
        // - average
        // -max
        Map<LocalDateTime, DoubleSummaryStatistics> mapAppelParJour;
        mapAppelParJour = this.callLogItemList.stream()
                .collect(Collectors.groupingBy(
                        e -> e.getDAYFromLocalDateTime(),
                        Collectors.summarizingDouble(e -> e.getDuration())
                ));
        LOG.log(Level.INFO, "Nombre d'éléments : " + mapAppelParJour.size());

        // Transformation de la Map en ExcelExcelRow 
        DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy/MM/dd");

        TreeMap<LocalDateTime, DoubleSummaryStatistics> treeMapAppelParJour;
        treeMapAppelParJour = new TreeMap<>(mapAppelParJour);

        treeMapAppelParJour.forEach(new BiConsumer<LocalDateTime, DoubleSummaryStatistics>() {
            int i = 1;

            @Override
            public void accept(LocalDateTime k, DoubleSummaryStatistics v) {

                Row row = sheet.createRow(i++);
                row.createCell(0).setCellValue(dtf.format(k));
                row.createCell(1).setCellValue(v.getCount());
                row.createCell(2).setCellValue(v.getSum() / 60);
                row.createCell(3).setCellValue(v.getAverage() / 60);
                row.createCell(4).setCellValue(v.getMin() / 60);
                row.createCell(5).setCellValue(v.getMax() / 60);

            }
        });
    }

    /**
     * Statistique des appels entrants par jour.
     *
     * @return Une description XML d'un ExcelWorksheet EXCEL au format XML.
     */
    private void worksheetAppelsParJourEntrantXlsx(Workbook wb) {
        String ongletName = "Appels entrants par jour";
        LOG.log(Level.INFO, "Création de l'onglet \"" + ongletName + "\"");

        short rowCpt = 0;

        CreationHelper creationHelper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet(ongletName);

        // Titre de colonnes
        //TODO Appliquer un style et une mise en forme à la 1ère ligne 
        Row row_0 = sheet.createRow(rowCpt);

        row_0.createCell(0).setCellValue("Date");
        row_0.createCell(1).setCellValue("Nombre d'appels entrants");
        row_0.createCell(2).setCellValue("Durée d'appels en minutes");
        row_0.createCell(3).setCellValue("Moyenne temps d'appels en minutes");
        row_0.createCell(4).setCellValue("Min temps d'appels en minutes");
        row_0.createCell(5).setCellValue("Max temps d'appels en minutes");

        // On enlève tous les appels sortants.
        List<CallLogItem> listAppelParJourFiltre;
        listAppelParJourFiltre = this.callLogItemList.stream()
                .filter(c -> !c.isOUTGOING_TYPE())
                .collect(Collectors.toList());

        // Statistiques par jour
        // Un objet DoubleSummaryStatistics contient 
        // - count
        // - sum
        // - min
        // - average
        // -max
        Map<LocalDateTime, DoubleSummaryStatistics> mapAppelParJour;
        mapAppelParJour = listAppelParJourFiltre.stream()
                .collect(Collectors.groupingBy(
                        e -> e.getDAYFromLocalDateTime(),
                        Collectors.summarizingDouble(e -> e.getDuration())
                ));
        LOG.log(Level.INFO, "Nombre d'éléments : " + mapAppelParJour.size());

        // Transformation de la Map en Worksheet 
        DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy/MM/dd");

        TreeMap<LocalDateTime, DoubleSummaryStatistics> treeMapAppelParJour;
        treeMapAppelParJour = new TreeMap<>(mapAppelParJour);

        treeMapAppelParJour.forEach(new BiConsumer<LocalDateTime, DoubleSummaryStatistics>() {
            int i = 1;

            @Override
            public void accept(LocalDateTime k, DoubleSummaryStatistics v) {

                Row row = sheet.createRow(i++);
                row.createCell(0).setCellValue(dtf.format(k));
                row.createCell(1).setCellValue(v.getCount());
                row.createCell(2).setCellValue(v.getSum() / 60);
                row.createCell(3).setCellValue(v.getAverage() / 60);
                row.createCell(4).setCellValue(v.getMin() / 60);
                row.createCell(5).setCellValue(v.getMax() / 60);

            }
        });

    }

    /**
     * Statistique des appels entrants par jour.
     *
     * @return Une description XML d'un ExcelWorksheet EXCEL au format XML.
     */
    private void worksheetAppelsParJourSortantXlsx(Workbook wb) {
        String ongletName = "Appels sortants par jour";
        LOG.log(Level.INFO, "Création de l'onglet \"" + ongletName + "\"");

        short rowCpt = 0;

        CreationHelper creationHelper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet(ongletName);

        // Titre de colonnes
        //TODO Appliquer un style et une mise en forme à la 1ère ligne 
        Row row_0 = sheet.createRow(rowCpt);

        row_0.createCell(0).setCellValue("Date");
        row_0.createCell(1).setCellValue("Nombre d'appels sortants");
        row_0.createCell(2).setCellValue("Durée d'appels en minutes");
        row_0.createCell(3).setCellValue("Moyenne temps d'appels en minutes");
        row_0.createCell(4).setCellValue("Min temps d'appels en minutes");
        row_0.createCell(5).setCellValue("Max temps d'appels en minutes");

        // On enlève tous les appels sortants.
        List<CallLogItem> listAppelParJourFiltre;
        listAppelParJourFiltre = this.callLogItemList.stream()
                .filter(c -> c.isOUTGOING_TYPE())
                .collect(Collectors.toList());

        // Statistiques par jour
        // Un objet DoubleSummaryStatistics contient 
        // - count
        // - sum
        // - min
        // - average
        // -max
        Map<LocalDateTime, DoubleSummaryStatistics> mapAppelParJour;
        mapAppelParJour = listAppelParJourFiltre.stream()
                .collect(Collectors.groupingBy(
                        e -> e.getDAYFromLocalDateTime(),
                        Collectors.summarizingDouble(e -> e.getDuration())
                ));
        LOG.log(Level.INFO, "Nombre d'éléments : " + mapAppelParJour.size());

        // Transformation de la Map en Worksheet 
        DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy/MM/dd");

        TreeMap<LocalDateTime, DoubleSummaryStatistics> treeMapAppelParJour;
        treeMapAppelParJour = new TreeMap<>(mapAppelParJour);

        treeMapAppelParJour.forEach(new BiConsumer<LocalDateTime, DoubleSummaryStatistics>() {
            int i = 1;

            @Override
            public void accept(LocalDateTime k, DoubleSummaryStatistics v) {

                Row row = sheet.createRow(i++);
                row.createCell(0).setCellValue(dtf.format(k));
                row.createCell(1).setCellValue(v.getCount());
                row.createCell(2).setCellValue(v.getSum() / 60);
                row.createCell(3).setCellValue(v.getAverage() / 60);
                row.createCell(4).setCellValue(v.getMin() / 60);
                row.createCell(5).setCellValue(v.getMax() / 60);

            }
        });

    }

    /**
     * Statistique des appels par mois.
     *
     * @return
     */
    private void worksheetAppelsParMoisXlsx(Workbook wb) {
        String ongletName = "Appels par mois";

        LOG.log(Level.INFO, "Création de l'onglet \"" + ongletName + "\"");

        short rowCpt = 0;

        CreationHelper creationHelper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet(ongletName);

        // Titre de colonnes 
        //TODO Appliquer un style et une mise en forme à la 1ère ligne 
        Row row_0 = sheet.createRow(rowCpt);

        row_0.createCell(0).setCellValue("Date");
        row_0.createCell(1).setCellValue("Nombre d'appels entrants");
        row_0.createCell(2).setCellValue("Durée d'appels en minutes");
        row_0.createCell(3).setCellValue("Moyenne temps d'appels en minutes");
        row_0.createCell(4).setCellValue("Min temps d'appels en minutes");
        row_0.createCell(5).setCellValue("Max temps d'appels en minutes");

        Map<String, DoubleSummaryStatistics> mapAppelParMois;
        mapAppelParMois = this.callLogItemList.stream()
                .collect(Collectors.groupingBy(
                        e -> e.getMONTHFromLocalDateTime(),
                        Collectors.summarizingDouble(e -> e.getDuration())
                ));
        LOG.log(Level.INFO, "Nombre d'éléments : " + mapAppelParMois.size());

        // Transformation de la Map en ExcelRow 
        DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy/MM/dd");

        TreeMap<String, DoubleSummaryStatistics> treeMapAppelParMois;
        treeMapAppelParMois = new TreeMap<>(mapAppelParMois);

        treeMapAppelParMois.forEach(new BiConsumer<String, DoubleSummaryStatistics>() {
            int i = 1;

            @Override
            public void accept(String k, DoubleSummaryStatistics v) {

                Row row = sheet.createRow(i++);
                row.createCell(0).setCellValue(k);
                row.createCell(1).setCellValue(v.getCount());
                row.createCell(2).setCellValue(v.getSum() / 60);
                row.createCell(3).setCellValue(v.getAverage() / 60);
                row.createCell(4).setCellValue(v.getMin() / 60);
                row.createCell(5).setCellValue(v.getMax() / 60);

            }
        });
    }

    /**
     * Statistique des appels entrants par mois.
     *
     */
    private void worksheetAppelsParMoisEntrantXlsx(Workbook wb) {
        String ongletName = "Appels entrant par mois";
        LOG.log(Level.INFO, "Création de l'onglet \"" + ongletName + "\"");

        short rowCpt = 0;

        CreationHelper creationHelper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet(ongletName);

        // Titre de colonnes 
        //TODO Appliquer un style et une mise en forme à la 1ère ligne 
        Row row_0 = sheet.createRow(rowCpt);

        row_0.createCell(0).setCellValue("Date");
        row_0.createCell(1).setCellValue("Nombre d'appels entrants");
        row_0.createCell(2).setCellValue("Durée d'appels en minutes");
        row_0.createCell(3).setCellValue("Moyenne temps d'appels en minutes");
        row_0.createCell(4).setCellValue("Min temps d'appels en minutes");
        row_0.createCell(5).setCellValue("Max temps d'appels en minutes");

        // On enlève tous les appels sortants.
        List<CallLogItem> listAppelParMoisFiltre;
        listAppelParMoisFiltre = this.callLogItemList.stream()
                .filter(c -> !c.isOUTGOING_TYPE())
                .collect(Collectors.toList());

        // Statistiques par jour
        Map<String, DoubleSummaryStatistics> mapAppelParMois;
        mapAppelParMois = listAppelParMoisFiltre.stream()
                .collect(Collectors.groupingBy(
                        e -> e.getMONTHFromLocalDateTime(),
                        Collectors.summarizingDouble(e -> e.getDuration())
                ));
        LOG.log(Level.INFO, "Nombre d'éléments : " + mapAppelParMois.size());

        // Transformation de la Map en ExcelRow 
        DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy/MM/dd");

        TreeMap<String, DoubleSummaryStatistics> treeMapAppelParMois;
        treeMapAppelParMois = new TreeMap<>(mapAppelParMois);

        treeMapAppelParMois.forEach(new BiConsumer<String, DoubleSummaryStatistics>() {
            int i = 1;

            @Override
            public void accept(String k, DoubleSummaryStatistics v) {

                Row row = sheet.createRow(i++);
                row.createCell(0).setCellValue(k);
                row.createCell(1).setCellValue(v.getCount());
                row.createCell(2).setCellValue(v.getSum() / 60);
                row.createCell(3).setCellValue(v.getAverage() / 60);
                row.createCell(4).setCellValue(v.getMin() / 60);
                row.createCell(5).setCellValue(v.getMax() / 60);

            }
        });
    }

    /**
     * Statistique des appels entrants par mois.
     *
     */
    private void worksheetAppelsParMoisSortantXlsx(Workbook wb) {
        String ongletName = "Appels sortants par mois";
        LOG.log(Level.INFO, "Création de l'onglet \"" + ongletName + "\"");

        short rowCpt = 0;

        CreationHelper creationHelper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet(ongletName);

        // Titre de colonnes 
        //TODO Appliquer un style et une mise en forme à la 1ère ligne 
        Row row_0 = sheet.createRow(rowCpt);

        row_0.createCell(0).setCellValue("Date");
        row_0.createCell(1).setCellValue("Nombre d'appels entrants");
        row_0.createCell(2).setCellValue("Durée d'appels en minutes");
        row_0.createCell(3).setCellValue("Moyenne temps d'appels en minutes");
        row_0.createCell(4).setCellValue("Min temps d'appels en minutes");
        row_0.createCell(5).setCellValue("Max temps d'appels en minutes");

        // On enlève tous les appels sortants.
        List<CallLogItem> listAppelParMoisFiltre;
        listAppelParMoisFiltre = this.callLogItemList.stream()
                .filter(c -> c.isOUTGOING_TYPE())
                .collect(Collectors.toList());

        // Statistiques par jour
        Map<String, DoubleSummaryStatistics> mapAppelParMois;
        mapAppelParMois = listAppelParMoisFiltre.stream()
                .collect(Collectors.groupingBy(
                        e -> e.getMONTHFromLocalDateTime(),
                        Collectors.summarizingDouble(e -> e.getDuration())
                ));
        LOG.log(Level.INFO, "Nombre d'éléments : " + mapAppelParMois.size());

        // Transformation de la Map en ExcelRow 
        DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy/MM/dd");

        TreeMap<String, DoubleSummaryStatistics> treeMapAppelParMois;
        treeMapAppelParMois = new TreeMap<>(mapAppelParMois);

        treeMapAppelParMois.forEach(new BiConsumer<String, DoubleSummaryStatistics>() {
            int i = 1;

            @Override
            public void accept(String k, DoubleSummaryStatistics v) {

                Row row = sheet.createRow(i++);
                row.createCell(0).setCellValue(k);
                row.createCell(1).setCellValue(v.getCount());
                row.createCell(2).setCellValue(v.getSum() / 60);
                row.createCell(3).setCellValue(v.getAverage() / 60);
                row.createCell(4).setCellValue(v.getMin() / 60);
                row.createCell(5).setCellValue(v.getMax() / 60);

            }
        });
    }

    /**
     * Statistique des appels par an.
     *
     */
    private void worksheetAppelsParAnXlsx(Workbook wb) {
        String ongletName = "Appels par an";

        LOG.log(Level.INFO, "Création de l'onglet \"" + ongletName + "\"");

        short rowCpt = 0;

        CreationHelper creationHelper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet(ongletName);

        // Titre de colonnes 
        //TODO Appliquer un style et une mise en forme à la 1ère ligne 
        Row row_0 = sheet.createRow(rowCpt);

        row_0.createCell(0).setCellValue("Année");
        row_0.createCell(1).setCellValue("Nombre d'appels entrants");
        row_0.createCell(2).setCellValue("Durée d'appels en minutes");
        row_0.createCell(3).setCellValue("Moyenne temps d'appels en minutes");
        row_0.createCell(4).setCellValue("Min temps d'appels en minutes");
        row_0.createCell(5).setCellValue("Max temps d'appels en minutes");

        // Statistiques par jour
        Map<String, DoubleSummaryStatistics> mapAppelParAn;
        mapAppelParAn = this.callLogItemList.stream()
                .collect(Collectors.groupingBy(
                        e -> e.getYEARFromLocalDateTime(),
                        Collectors.summarizingDouble(e -> e.getDuration())
                ));
        LOG.log(Level.INFO, "Nombre d'éléments : " + mapAppelParAn.size());

        TreeMap<String, DoubleSummaryStatistics> treeMaprAppelParAn;
        treeMaprAppelParAn = new TreeMap<>(mapAppelParAn);

        treeMaprAppelParAn.forEach(new BiConsumer<String, DoubleSummaryStatistics>() {
            int i = 1;

            @Override
            public void accept(String k, DoubleSummaryStatistics v) {

                Row row = sheet.createRow(i++);
                row.createCell(0).setCellValue(k);
                row.createCell(1).setCellValue(v.getCount());
                row.createCell(2).setCellValue(v.getSum() / 60);
                row.createCell(3).setCellValue(v.getAverage() / 60);
                row.createCell(4).setCellValue(v.getMin() / 60);
                row.createCell(5).setCellValue(v.getMax() / 60);

            }
        });
    }

    /**
     * Statistiques des appels sortants par an.
     *
     */
    private void worksheetAppelsParAnEntrantXlsx(Workbook wb) {
        String ongletName = "Appels entrants par an";

        LOG.log(Level.INFO, "Création de l'onglet \"" + ongletName + "\"");

        short rowCpt = 0;

        CreationHelper creationHelper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet(ongletName);

        // Titre de colonnes 
        //TODO Appliquer un style et une mise en forme à la 1ère ligne 
        Row row_0 = sheet.createRow(rowCpt);

        row_0.createCell(0).setCellValue("Année");
        row_0.createCell(1).setCellValue("Nombre d'appels entrants");
        row_0.createCell(2).setCellValue("Durée d'appels en minutes");
        row_0.createCell(3).setCellValue("Moyenne temps d'appels en minutes");
        row_0.createCell(4).setCellValue("Min temps d'appels en minutes");
        row_0.createCell(5).setCellValue("Max temps d'appels en minutes");

        // Statistiques par jour
        // On enlève tous les appels sortants.
        List<CallLogItem> listAppelParMoisFiltre;
        listAppelParMoisFiltre = this.callLogItemList.stream()
                .filter(c -> !c.isOUTGOING_TYPE())
                .collect(Collectors.toList());

        Map<String, DoubleSummaryStatistics> mapAppelParAn;
        mapAppelParAn = listAppelParMoisFiltre.stream()
                .collect(Collectors.groupingBy(
                        e -> e.getYEARFromLocalDateTime(),
                        Collectors.summarizingDouble(e -> e.getDuration())
                ));
        LOG.log(Level.INFO, "Nombre d'éléments : " + mapAppelParAn.size());

        TreeMap<String, DoubleSummaryStatistics> treeMaprAppelParAn;
        treeMaprAppelParAn = new TreeMap<>(mapAppelParAn);

        treeMaprAppelParAn.forEach(new BiConsumer<String, DoubleSummaryStatistics>() {
            int i = 1;

            @Override
            public void accept(String k, DoubleSummaryStatistics v) {

                Row row = sheet.createRow(i++);
                row.createCell(0).setCellValue(k);
                row.createCell(1).setCellValue(v.getCount());
                row.createCell(2).setCellValue(v.getSum() / 60);
                row.createCell(3).setCellValue(v.getAverage() / 60);
                row.createCell(4).setCellValue(v.getMin() / 60);
                row.createCell(5).setCellValue(v.getMax() / 60);

            }
        });
    }

    /**
     * Statistiques des appels sortants par an.
     *
     */
    private void worksheetAppelsParAnSortantXlsx(Workbook wb) {
        String ongletName = "Appels sortants par an";

        LOG.log(Level.INFO, "Création de l'onglet \"" + ongletName + "\"");

        short rowCpt = 0;

        CreationHelper creationHelper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet(ongletName);

        // Titre de colonnes 
        //TODO Appliquer un style et une mise en forme à la 1ère ligne 
        Row row_0 = sheet.createRow(rowCpt);

        row_0.createCell(0).setCellValue("Année");
        row_0.createCell(1).setCellValue("Nombre d'appels entrants");
        row_0.createCell(2).setCellValue("Durée d'appels en minutes");
        row_0.createCell(3).setCellValue("Moyenne temps d'appels en minutes");
        row_0.createCell(4).setCellValue("Min temps d'appels en minutes");
        row_0.createCell(5).setCellValue("Max temps d'appels en minutes");

        // Statistiques par jour
        // On enlève tous les appels sortants.
        List<CallLogItem> listAppelParMoisFiltre;
        listAppelParMoisFiltre = this.callLogItemList.stream()
                .filter(c -> c.isOUTGOING_TYPE())
                .collect(Collectors.toList());

        Map<String, DoubleSummaryStatistics> mapAppelParAn;
        mapAppelParAn = listAppelParMoisFiltre.stream()
                .collect(Collectors.groupingBy(
                        e -> e.getYEARFromLocalDateTime(),
                        Collectors.summarizingDouble(e -> e.getDuration())
                ));
        LOG.log(Level.INFO, "Nombre d'éléments : " + mapAppelParAn.size());

        TreeMap<String, DoubleSummaryStatistics> treeMaprAppelParAn;
        treeMaprAppelParAn = new TreeMap<>(mapAppelParAn);

        treeMaprAppelParAn.forEach(new BiConsumer<String, DoubleSummaryStatistics>() {
            int i = 1;

            @Override
            public void accept(String k, DoubleSummaryStatistics v) {

                Row row = sheet.createRow(i++);
                row.createCell(0).setCellValue(k);
                row.createCell(1).setCellValue(v.getCount());
                row.createCell(2).setCellValue(v.getSum() / 60);
                row.createCell(3).setCellValue(v.getAverage() / 60);
                row.createCell(4).setCellValue(v.getMin() / 60);
                row.createCell(5).setCellValue(v.getMax() / 60);

            }
        });

    }

    /**
     * Statistique des appels par tranche de 24H
     *
     */
    private void worksheetAppelsPar24HXlsx(Workbook wb) {
        String ongletName = "Appels par tranche de 24h";

        LOG.log(Level.INFO, "Création de l'onglet \"" + ongletName + "\"");

        short rowCpt = 0;

        CreationHelper creationHelper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet(ongletName);

        // Titre de colonnes 
        //TODO Appliquer un style et une mise en forme à la 1ère ligne 
        Row row_0 = sheet.createRow(rowCpt);

        row_0.createCell(0).setCellValue("Tranche de 24h");
        row_0.createCell(1).setCellValue("Nombre d'appels entrants");
        row_0.createCell(2).setCellValue("Durée d'appels en minutes");
        row_0.createCell(3).setCellValue("Moyenne temps d'appels en minutes");
        row_0.createCell(4).setCellValue("Min temps d'appels en minutes");
        row_0.createCell(5).setCellValue("Max temps d'appels en minutes");

        // Statistiques par jour
        Map<String, DoubleSummaryStatistics> mapAppelParAn;
        mapAppelParAn = this.callLogItemList.stream()
                .collect(Collectors.groupingBy(
                        e -> e.getHOURSFromLocalDateTime(),
                        Collectors.summarizingDouble(e -> e.getDuration())
                ));
        LOG.log(Level.INFO, "Nombre d'éléments : " + mapAppelParAn.size());

        TreeMap<String, DoubleSummaryStatistics> treeMaprAppelParAn;
        treeMaprAppelParAn = new TreeMap<>(mapAppelParAn);

        treeMaprAppelParAn.forEach(new BiConsumer<String, DoubleSummaryStatistics>() {
            int i = 1;

            @Override
            public void accept(String k, DoubleSummaryStatistics v) {

                Row row = sheet.createRow(i++);
                row.createCell(0).setCellValue(k);
                row.createCell(1).setCellValue(v.getCount());
                row.createCell(2).setCellValue(v.getSum() / 60);
                row.createCell(3).setCellValue(v.getAverage() / 60);
                row.createCell(4).setCellValue(v.getMin() / 60);
                row.createCell(5).setCellValue(v.getMax() / 60);

            }
        });
    }

    /**
     * Statistique des appels entrants par tranche de 24H
     *
     */
    private void worksheetAppelsPar24HEntrantsXlsx(Workbook wb) {
        String ongletName = "Appels entrants par tranche de 24h";

        LOG.log(Level.INFO, "Création de l'onglet \"" + ongletName + "\"");

        short rowCpt = 0;

        CreationHelper creationHelper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet(ongletName);

        // Titre de colonnes 
        //TODO Appliquer un style et une mise en forme à la 1ère ligne 
        Row row_0 = sheet.createRow(rowCpt);

        row_0.createCell(0).setCellValue("Tranche de 24h");
        row_0.createCell(1).setCellValue("Nombre d'appels entrants");
        row_0.createCell(2).setCellValue("Durée d'appels en minutes");
        row_0.createCell(3).setCellValue("Moyenne temps d'appels en minutes");
        row_0.createCell(4).setCellValue("Min temps d'appels en minutes");
        row_0.createCell(5).setCellValue("Max temps d'appels en minutes");

        // On enlève tous les appels sortants.
        List<CallLogItem> listAppelParJourFiltre;
        listAppelParJourFiltre = this.callLogItemList.stream()
                .filter(c -> !c.isOUTGOING_TYPE())
                .collect(Collectors.toList());

        // Statistiques par jour
        Map<String, DoubleSummaryStatistics> mapAppelPar24H;
        mapAppelPar24H = listAppelParJourFiltre.stream()
                .collect(Collectors.groupingBy(
                        e -> e.getHOURSFromLocalDateTime(),
                        Collectors.summarizingDouble(e -> e.getDuration())
                ));
        LOG.log(Level.INFO, "Nombre d'éléments : " + mapAppelPar24H.size());

        TreeMap<String, DoubleSummaryStatistics> treeMaprAppelParAn;
        treeMaprAppelParAn = new TreeMap<>(mapAppelPar24H);

        treeMaprAppelParAn.forEach(new BiConsumer<String, DoubleSummaryStatistics>() {
            int i = 1;

            @Override
            public void accept(String k, DoubleSummaryStatistics v) {

                Row row = sheet.createRow(i++);
                row.createCell(0).setCellValue(k);
                row.createCell(1).setCellValue(v.getCount());
                row.createCell(2).setCellValue(v.getSum() / 60);
                row.createCell(3).setCellValue(v.getAverage() / 60);
                row.createCell(4).setCellValue(v.getMin() / 60);
                row.createCell(5).setCellValue(v.getMax() / 60);

            }
        });
    }

    /**
     * Statistique des appels entrants par tranche de 24H
     *
     */
    private void worksheetAppelsPar24HSortantsXlsx(Workbook wb) {
        String ongletName = "Appels sortants par tranche de 24h";

        LOG.log(Level.INFO, "Création de l'onglet \"" + ongletName + "\"");

        short rowCpt = 0;

        CreationHelper creationHelper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet(ongletName);

        // Titre de colonnes 
        //TODO Appliquer un style et une mise en forme à la 1ère ligne 
        Row row_0 = sheet.createRow(rowCpt);

        row_0.createCell(0).setCellValue("Tranche de 24h");
        row_0.createCell(1).setCellValue("Nombre d'appels entrants");
        row_0.createCell(2).setCellValue("Durée d'appels en minutes");
        row_0.createCell(3).setCellValue("Moyenne temps d'appels en minutes");
        row_0.createCell(4).setCellValue("Min temps d'appels en minutes");
        row_0.createCell(5).setCellValue("Max temps d'appels en minutes");

        // On enlève tous les appels sortants.
        List<CallLogItem> listAppelParJourFiltre;
        listAppelParJourFiltre = this.callLogItemList.stream()
                .filter(c -> c.isOUTGOING_TYPE())
                .collect(Collectors.toList());

        // Statistiques par jour
        Map<String, DoubleSummaryStatistics> mapAppelPar24H;
        mapAppelPar24H = listAppelParJourFiltre.stream()
                .collect(Collectors.groupingBy(
                        e -> e.getHOURSFromLocalDateTime(),
                        Collectors.summarizingDouble(e -> e.getDuration())
                ));
        LOG.log(Level.INFO, "Nombre d'éléments : " + mapAppelPar24H.size());

        TreeMap<String, DoubleSummaryStatistics> treeMaprAppelParAn;
        treeMaprAppelParAn = new TreeMap<>(mapAppelPar24H);

        treeMaprAppelParAn.forEach(new BiConsumer<String, DoubleSummaryStatistics>() {
            int i = 1;

            @Override
            public void accept(String k, DoubleSummaryStatistics v) {

                Row row = sheet.createRow(i++);
                row.createCell(0).setCellValue(k);
                row.createCell(1).setCellValue(v.getCount());
                row.createCell(2).setCellValue(v.getSum() / 60);
                row.createCell(3).setCellValue(v.getAverage() / 60);
                row.createCell(4).setCellValue(v.getMin() / 60);
                row.createCell(5).setCellValue(v.getMax() / 60);

            }
        });
    }

    /**
     * Statistique des appels par jour de la semaine
     *
     */
    private void worksheetAppelsParJourDeLaSemaineXlsx(Workbook wb) {
        String ongletName = "Appels par jour de la semaine";

        LOG.log(Level.INFO, "Création de l'onglet \"" + ongletName + "\"");

        short rowCpt = 0;

        CreationHelper creationHelper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet(ongletName);

        // Titre de colonnes 
        //TODO Appliquer un style et une mise en forme à la 1ère ligne 
        Row row_0 = sheet.createRow(rowCpt);

        row_0.createCell(0).setCellValue("Jour de la semaine");
        row_0.createCell(1).setCellValue("Nombre d'appels entrants");
        row_0.createCell(2).setCellValue("Durée d'appels en minutes");
        row_0.createCell(3).setCellValue("Moyenne temps d'appels en minutes");
        row_0.createCell(4).setCellValue("Min temps d'appels en minutes");
        row_0.createCell(5).setCellValue("Max temps d'appels en minutes");

        // Statistiques par jour
        Map<DayOfWeek, DoubleSummaryStatistics> mapAppelParAn;
        mapAppelParAn = this.callLogItemList.stream()
                .collect(Collectors.groupingBy(
                        e -> e.getDayOfWeekFromLocalDateTime(),
                        Collectors.summarizingDouble(e -> e.getDuration())
                ));
        LOG.log(Level.INFO, "Nombre d'éléments : " + mapAppelParAn.size());

        TreeMap<DayOfWeek, DoubleSummaryStatistics> treeMaprAppelParJourDeLaSemaine;
        treeMaprAppelParJourDeLaSemaine = new TreeMap<>(mapAppelParAn);

        treeMaprAppelParJourDeLaSemaine.forEach(new BiConsumer<DayOfWeek, DoubleSummaryStatistics>() {
            int i = 1;

            @Override
            public void accept(DayOfWeek k, DoubleSummaryStatistics v) {

                Row row = sheet.createRow(i++);
                row.createCell(0).setCellValue(k.toString());
                row.createCell(1).setCellValue(v.getCount());
                row.createCell(2).setCellValue(v.getSum() / 60);
                row.createCell(3).setCellValue(v.getAverage() / 60);
                row.createCell(4).setCellValue(v.getMin() / 60);
                row.createCell(5).setCellValue(v.getMax() / 60);

            }
        });
    }

    /**
     * Statistique des appels par jour de la semaine
     *
     */
    private void worksheetAppelsParJourDeLaSemaineEntrantsXlsx(Workbook wb) {
        String ongletName = "Appels entrants par jour de la semaine";

        LOG.log(Level.INFO, "Création de l'onglet \"" + ongletName + "\"");

        short rowCpt = 0;

        CreationHelper creationHelper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet(ongletName);

        // Titre de colonnes 
        //TODO Appliquer un style et une mise en forme à la 1ère ligne 
        Row row_0 = sheet.createRow(rowCpt);

        row_0.createCell(0).setCellValue("Jour de la semaine");
        row_0.createCell(1).setCellValue("Nombre d'appels entrants");
        row_0.createCell(2).setCellValue("Durée d'appels en minutes");
        row_0.createCell(3).setCellValue("Moyenne temps d'appels en minutes");
        row_0.createCell(4).setCellValue("Min temps d'appels en minutes");
        row_0.createCell(5).setCellValue("Max temps d'appels en minutes");

        // On enlève tous les appels sortants.
        List<CallLogItem> listAppelParJourDeLaSemainesFiltre;
        listAppelParJourDeLaSemainesFiltre = this.callLogItemList.stream()
                .filter(c -> !c.isOUTGOING_TYPE())
                .collect(Collectors.toList());

        // Statistiques par jour
        Map<DayOfWeek, DoubleSummaryStatistics> mapAppelParJourDeLaSemaine;
        mapAppelParJourDeLaSemaine = listAppelParJourDeLaSemainesFiltre.stream()
                .collect(Collectors.groupingBy(
                        e -> e.getDayOfWeekFromLocalDateTime(),
                        Collectors.summarizingDouble(e -> e.getDuration())
                ));
        LOG.log(Level.INFO, "Nombre d'éléments : " + mapAppelParJourDeLaSemaine.size());

        TreeMap<DayOfWeek, DoubleSummaryStatistics> treeMaprAppelParAn;
        treeMaprAppelParAn = new TreeMap<>(mapAppelParJourDeLaSemaine);

        treeMaprAppelParAn.forEach(new BiConsumer<DayOfWeek, DoubleSummaryStatistics>() {
            int i = 1;

            @Override
            public void accept(DayOfWeek k, DoubleSummaryStatistics v) {

                Row row = sheet.createRow(i++);
                row.createCell(0).setCellValue(k.toString());
                row.createCell(1).setCellValue(v.getCount());
                row.createCell(2).setCellValue(v.getSum() / 60);
                row.createCell(3).setCellValue(v.getAverage() / 60);
                row.createCell(4).setCellValue(v.getMin() / 60);
                row.createCell(5).setCellValue(v.getMax() / 60);

            }
        });
    }

    /**
     * Statistique des appels par jour de la semaine
     *
     */
    private void worksheetAppelsParJourDeLaSemaineSortantsXlsx(Workbook wb) {
        String ongletName = "Appels sortants par jour de la semaine";

        LOG.log(Level.INFO, "Création de l'onglet \"" + ongletName + "\"");

        short rowCpt = 0;

        CreationHelper creationHelper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet(ongletName);

        // Titre de colonnes 
        //TODO Appliquer un style et une mise en forme à la 1ère ligne 
        Row row_0 = sheet.createRow(rowCpt);

        row_0.createCell(0).setCellValue("Jour de la semaine");
        row_0.createCell(1).setCellValue("Nombre d'appels entrants");
        row_0.createCell(2).setCellValue("Durée d'appels en minutes");
        row_0.createCell(3).setCellValue("Moyenne temps d'appels en minutes");
        row_0.createCell(4).setCellValue("Min temps d'appels en minutes");
        row_0.createCell(5).setCellValue("Max temps d'appels en minutes");

        // On enlève tous les appels sortants.
        List<CallLogItem> listAppelParJourDeLaSemainesFiltre;
        listAppelParJourDeLaSemainesFiltre = this.callLogItemList.stream()
                .filter(c -> !c.isOUTGOING_TYPE())
                .collect(Collectors.toList());

        // Statistiques par jour
        Map<DayOfWeek, DoubleSummaryStatistics> mapAppelParJourDeLaSemaine;
        mapAppelParJourDeLaSemaine = listAppelParJourDeLaSemainesFiltre.stream()
                .collect(Collectors.groupingBy(
                        e -> e.getDayOfWeekFromLocalDateTime(),
                        Collectors.summarizingDouble(e -> e.getDuration())
                ));
        LOG.log(Level.INFO, "Nombre d'éléments : " + mapAppelParJourDeLaSemaine.size());

        TreeMap<DayOfWeek, DoubleSummaryStatistics> treeMaprAppelParAn;
        treeMaprAppelParAn = new TreeMap<>(mapAppelParJourDeLaSemaine);

        treeMaprAppelParAn.forEach(new BiConsumer<DayOfWeek, DoubleSummaryStatistics>() {
            int i = 1;

            @Override
            public void accept(DayOfWeek k, DoubleSummaryStatistics v) {

                Row row = sheet.createRow(i++);
                row.createCell(0).setCellValue(k.toString());
                row.createCell(1).setCellValue(v.getCount());
                row.createCell(2).setCellValue(v.getSum() / 60);
                row.createCell(3).setCellValue(v.getAverage() / 60);
                row.createCell(4).setCellValue(v.getMin() / 60);
                row.createCell(5).setCellValue(v.getMax() / 60);

            }
        });
    }
}
