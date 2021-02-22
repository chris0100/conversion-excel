package com.feho.cambiofecha.services;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;

public interface FechaService {

    // Convertir la fecha
    List<String> convertirFecha() throws Exception;

    //convertir a excel
    void convertStringToExcel() throws IOException, InvalidFormatException;

    // convertir a csv
    void crearArchivoCSV(String nombreArchivo);
}
