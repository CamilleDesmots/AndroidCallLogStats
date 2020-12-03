/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package org.camilledesmots.androidcalllogstats;

import java.time.Instant;
import java.time.DayOfWeek;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoUnit;

/**
 *
 * @author camilledesmots
 */
public class CallLogItem {

    private static final int INCOMING_TYPE = 1;
    private static final int OUTGOING_TYPE = 2;
    private static final int MISSED_TYPE = 3;

    private final DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:MM");
    
    /**
     * The phone number
     */
    private String number;

    /**
     * The type of the call (incoming, outgoing or missed).
     */
    private Integer type;

    /**
     * The time the call occured, in milliseconds since the epoch
     */
//    private Long epoch;
    /**
     * The time the call occured, in milliseconds since the epoch
     */
    private LocalDateTime localDateTime;

    /**
     * The duration of the call in seconds
     */
    private Long duration;

    /**
     * Don't know what is this column 5 content yet
     */
    private Integer column5;

    /**
     * Don't know what is this column 6 content yet
     */
    private Integer column6;

    public String getNumber() {
        return number;
    }

    public void setNumber(String number) {
        this.number = number;
    }

    public Integer getType() {
        return type;
    }

    public void setType(Integer type) {
        this.type = type;
    }

//    public Long getEpoch() {
//        return epoch;
//    }
//    public void setEpoch(Long epoch) {
//        this.epoch = epoch;
//    }
    public LocalDateTime getLocalDateTime() {
        return localDateTime;
    }

    public void setLocalDateTime(LocalDateTime localDateTime) {
        this.localDateTime = localDateTime;
    }

    public void setLocalDateTime(long epoch) {
        this.localDateTime = Instant.ofEpochMilli(epoch)
                .atZone(ZoneId.systemDefault()).toLocalDateTime();
    }

    public Long getDuration() {
        return duration;
    }

    public void setDuration(Long duration) {
        this.duration = duration;
    }

    public Integer getColumn5() {
        return column5;
    }

    public void setColumn5(Integer column5) {
        this.column5 = column5;
    }

    public Integer getColumn6() {
        return column6;
    }

    public void setColumn6(Integer column6) {
        this.column6 = column6;
    }

    public String getHOURSFromLocalDateTime() {
        return this.localDateTime.toString().substring(11, 13);
    }
    
       public DayOfWeek getDayOfWeekFromLocalDateTime() {
        return this.localDateTime.getDayOfWeek();
    }


    public LocalDateTime getDAYFromLocalDateTime() {
        return this.localDateTime.truncatedTo(ChronoUnit.DAYS);
    }

    public String getMONTHFromLocalDateTime() {
        LocalDateTime ldt;
        ldt = this.localDateTime.truncatedTo(ChronoUnit.DAYS);
        dtf.format(ldt);
        return ldt.toString().substring(0, 7);
    }

    public String getYEARFromLocalDateTime() {
        LocalDateTime ldt;
        ldt = this.localDateTime.truncatedTo(ChronoUnit.DAYS);
        dtf.format(ldt);
        return ldt.toString().substring(0, 4);
    }

    /**
     * Constructeur
     *
     * @param number Le numéro de téléphone
     * @param type Le type d'appel
     * @param time La time d'appel en millisecondes depuis le 1er janvier 1960
     * @param duration La durée d'appel en millissecondes
     */
    @SuppressWarnings("empty-statement")
    public void CallLogItem(String number, Integer type, long epoch, Long duration) {
        this.number = number;
        this.type = type;
        this.localDateTime = Instant.ofEpochMilli(epoch)
                .atZone(ZoneId.systemDefault()).toLocalDateTime();
        this.duration = duration;
        this.column5 = 0;
        this.column6 = 0;
    }

    /**
     * Appel entrant ?
     *
     * @return TRUE Vrai
     */
    public Boolean isINCOMING_TYPE() {
        return this.type == INCOMING_TYPE;
    }

    /**
     * Appel sortant ?
     *
     * @return TRUE Vrai
     */
    public Boolean isOUTGOING_TYPE() {
        return this.type == OUTGOING_TYPE;
    }

    /**
     * Appel manqué ?
     *
     * @return TRUE Vrai
     */
    public Boolean isMISSED_TYPE() {
        return this.type == MISSED_TYPE;
    }

    public String toString() {
        return ("number = " + this.number
                + " date = " + this.localDateTime
                + " type = " + this.type
                + " duration = " + this.duration);
    }

}
