package com.mkdika.poihelper.demo;

import java.util.Date;

public class Person {

    private Integer id;
    private String fullName;
    private Date birthDate;
    private Double ballance;

    public Person(){}

    public Person(Integer id, String fullName, Date birthDate, Double ballance) {
        this.id = id;
        this.fullName = fullName;
        this.birthDate = birthDate;
        this.ballance = ballance;
    }

    public Integer getId() {
        return id;
    }

    public void setId(Integer id) {
        this.id = id;
    }

    public String getFullName() {
        return fullName;
    }

    public void setFullName(String fullName) {
        this.fullName = fullName;
    }

    public Date getBirthDate() {
        return birthDate;
    }

    public void setBirthDate(Date birthDate) {
        this.birthDate = birthDate;
    }

    public Double getBallance() {
        return ballance;
    }

    public void setBallance(Double ballance) {
        this.ballance = ballance;
    }
}
