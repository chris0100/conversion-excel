package com.feho.cambiofecha.entities;


import lombok.*;

import java.util.Date;

@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
@ToString
public class Employee {

    private String name;
    private String email;
    private Date dateOfBirth;
    private double salary;
}
