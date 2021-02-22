package com.feho.cambiofecha.services;


import com.feho.cambiofecha.entities.Employee;
import lombok.extern.log4j.Log4j2;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Service;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;


@Log4j2
@Service
public class FechaServiceImpl implements FechaService {

    private static String[] columns = {"Name", "Email", "Date Of Birth", "Salary"};
    private static List<Employee> employees = new ArrayList<>();


    @Override
    public List<String> convertirFecha() throws Exception {

        /*String zona1 = "America/Bogota";

        ZonedDateTime fechaInventada = ZonedDateTime.of(2021, 2, 4, 12, 00, 00, 00, ZoneId.of("Europe/Madrid"));

        System.out.println("********");
        System.out.println(fechaInventada);
        System.out.println("********");

        // Se inserta la fecha con la hora actual
        Date fechaActual = new Date();

        // Se Obtiene el zonedDateTime de mi ubicacion actual
        ZonedDateTime zonaActualCar = ZonedDateTime.ofInstant(fechaActual.toInstant(), ZoneId.systemDefault());

        // Se obtiene el zonedDateTime de la ubicacion de mi preferencia
        ZonedDateTime zonaPreferida = fechaInventada.withZoneSameInstant(ZoneId.of(zona1));

        // Se convierte a DateTimeFormat
        LocalDateTime convertida = zonaPreferida.toLocalDateTime();
        System.out.println(convertida);
        DateTimeFormatter dtf = DateTimeFormatter.ofPattern("uuuuMMddkkmmss");
        System.out.println(dtf.format(convertida));


        // Se crea una lista para añadir la fecha
        List<String> zonasHorarias = new ArrayList<>();

        //Se carga la zona1
        zonasHorarias.add(dtf.format(convertida));*/


        System.out.println(dateToMilliseconds("202102040000"));

        List<String> zonasHorarias = new ArrayList<>();

        return zonasHorarias;
    }


    private long dateToMilliseconds(String date) {
        try {
            //convertir a date
            Date date1 = new SimpleDateFormat("yyyyMMddhhmm").parse(date);
            System.out.println("*********convertido");
            System.out.println(date1);

            String zone = "America/Bogota";


            // zonedDateTime from actual location
            ZonedDateTime actualZone = ZonedDateTime.ofInstant(date1.toInstant(), ZoneId.systemDefault());

            // zonedDateTime to prefer location
            ZonedDateTime preferZone = actualZone.withZoneSameInstant(ZoneId.of(zone));

            // convert the LocalDateTime to format
            LocalDateTime converted = preferZone.toLocalDateTime();
            System.out.println(converted);
            DateTimeFormatter dtf = DateTimeFormatter.ofPattern("uuuuMMddkkmm");
            System.out.println(dtf.format(converted));

            String nuevoDate = dtf.format(converted);

            List<SimpleDateFormat> formatsDateList = new ArrayList<>();
            formatsDateList.add(new SimpleDateFormat("yyyyMMddhhmm"));
            formatsDateList.add(new SimpleDateFormat("yyyyMMdd"));


            for (SimpleDateFormat pattern : formatsDateList) {
                try {
                    return pattern.parse(nuevoDate).getTime();
                } catch (ParseException ex) {
                    System.out.println("error");
                }
            }
            return 0;
        } catch (Exception e) {
            System.out.println("error");
            return -1;
        }
    }


    //***************************************************************************************
    //***************************************************************************************
    //***************************************************************************************
    //***************************************************************************************
    //***************************************************************************************
    //***************************************************************************************
    //***************************************************************************************

    //convertir a excel
    @Override
    public void convertStringToExcel() throws IOException, InvalidFormatException {


        // Añadir datos a la lista de empleados
        Calendar dateOfBirth = Calendar.getInstance();
        dateOfBirth.set(1992, 7, 21);
        employees.add(new Employee("Rajeev Singh", "rajeev@example.com",
                dateOfBirth.getTime(), 1200000.0));

        dateOfBirth.set(1965, 10, 15);
        employees.add(new Employee("Thomas cook", "thomas@example.com",
                dateOfBirth.getTime(), 1500000.0));

        dateOfBirth.set(1987, 4, 18);
        employees.add(new Employee("Steve Maiden", "steve@example.com",
                dateOfBirth.getTime(), 1800000.0));


        // Crear libro de excel
        Workbook workbook = new XSSFWorkbook();

        /* CreationHelper helps us create instances of various things like DataFormat,
           Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way */
        CreationHelper createHelper = workbook.getCreationHelper();

        // Crear la hoja
        Sheet sheet = workbook.createSheet("Employee");

        // Crear Estilos de fuente para los encabezados
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.RED.getIndex());

        // Crear un estilo de celda para la fuente
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        // Crear fila
        Row headerRow = sheet.createRow(0);

        // Crear celdas
        for (int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
        }

        // Crear estilo de celdas para formatear la fecha
        CellStyle dateCellStyle = workbook.createCellStyle();
        dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd-MM-yyyy"));

        // Rellenar los datos en las filas y columnas
        int rowNum = 1;
        for (Employee employee : employees) {
            Row row = sheet.createRow(rowNum++);

            row.createCell(0)
                    .setCellValue(employee.getName());

            row.createCell(1)
                    .setCellValue(employee.getEmail());

            Cell dateOfBirthCell = row.createCell(2);
            dateOfBirthCell.setCellValue(employee.getDateOfBirth());
            dateOfBirthCell.setCellStyle(dateCellStyle);

            row.createCell(3)
                    .setCellValue(employee.getSalary());
        }

        // ajusta el ancho de las columnas de acuerdo al contenido
        for (int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // Escribe la info en un archivo nuevo
        //FileOutputStream fileOut = new FileOutputStream("poi-generated-file.xlsx");
        //workbook.write(fileOut);


        try {
            //Se obtiene el bytearray de
            //byte[] convertidoBytes = Files.readAllBytes(Paths.get("poi-generated-file.xlsx"));

            // obtener el bytearray de ese archivo
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            workbook.write(baos);
            byte[] baosConverted = baos.toByteArray();

            // lo codifica en base 64
            String encoded = Base64.getEncoder().encodeToString(baosConverted);

            // Cierra el libro
            //fileOut.close();
            workbook.close();

            // envia el content(binary data) en base 64 string y lo arma en un archivo
            specialConversion(encoded);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    // Pasar de base64 a string y crear archivo de excel(Simulacion de lo que haria el otro microservicio cuando recibe el content base64)
    private void specialConversion(String base64cod) throws IOException {
        //Lo decodifica y lo vuelve byte array...
        byte[] decode = Base64.getDecoder().decode(base64cod);

        //escribir en un archivo nuevo
        String fileOutput = "salida2.xlsx";
        FileOutputStream fos = new FileOutputStream(fileOutput);
        fos.write(decode);
    }





    //ESTE METODO ES PARA PRUEBAS DE CONVERSIONES Y GENERACION DE ARCHIVO
    //Lee archivo para pasar a bytes[]
    private void pasaToBytes() {
        try {
            byte[] bytes = Files.readAllBytes(Paths.get("poi-generated-file.xlsx"));
            log.info("nueva conversion y obtiene bytes[]");
            log.info(bytes);

            //escribir
            String fileOutput = "salida.xlsx";
            FileOutputStream fos = new FileOutputStream(fileOutput);
            fos.write(bytes);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    // SEGUNDA OPCION QUE SE PLANEABA MANEJAR CON CSV
    // Convertir a CSV
    @Override
    public void crearArchivoCSV(String nombreDeArchivo) {
        // this also can be "\t" for tab delimitador
        escribeArchivoCSV(nombreDeArchivo, ";");
    }

    private static void escribeArchivoCSV(String file, String delim) {
        final String NEXT_LINE = "\n";
        try {
            FileWriter fw = new FileWriter(file);

            fw.append("testing").append(delim);
            fw.append("123").append(NEXT_LINE);

            fw.append("value1");
            fw.append(delim);
            fw.append("312");
            fw.append(NEXT_LINE);

            fw.flush();
            fw.close();
        } catch (IOException e) {
            // Error al crear el archivo, por ejemplo, el archivo
            // está actualmente abierto.
            e.printStackTrace();
        }
    }

}
