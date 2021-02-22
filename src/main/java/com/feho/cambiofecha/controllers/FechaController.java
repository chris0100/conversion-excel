package com.feho.cambiofecha.controllers;

import com.feho.cambiofecha.services.FechaService;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@RestController
public class FechaController {

    @Autowired
    private FechaService fechaServiceObj;


    // Crear archivo Excel
    @GetMapping("/excel")
    public ResponseEntity crearXLSX() throws IOException, InvalidFormatException {
        fechaServiceObj.convertStringToExcel();
        return new ResponseEntity<>("funciono excel", HttpStatus.OK);
    }


    // Crear archivo CSV
    @GetMapping("/csv")
    public ResponseEntity crearCSV(){
        // crear archivo csv
        final String nombreDeArchivo = "archivo.csv";
        fechaServiceObj.crearArchivoCSV(nombreDeArchivo);
        return new ResponseEntity<>("funciono csv", HttpStatus.OK);
    }


}
