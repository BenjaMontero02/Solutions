package com.bonartech.springbackend.Services.utils.reader.Reader;

import jakarta.servlet.ServletOutputStream;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;

public abstract class ReportAbstract {

    //Archivo .xlxs
    protected XSSFWorkbook workbook;
    //Hoja a trabajar
    protected XSSFSheet sheet;

    public ReportAbstract(){
        this.workbook = new XSSFWorkbook();
    }

    protected abstract void instancedSheet();

    //creo una celda
    //row es la fila
    //columnCount, posicion de la columna
    //valueOfCell valor a poner en la columna
    //Style estilo que se le asigna a la celda
    protected void createCell(Row row, int columnCount, Object valueOfCell, XSSFCellStyle style) {
        sheet.autoSizeColumn(columnCount);
        Cell cell = row.createCell(columnCount);
        if (valueOfCell instanceof Integer) {
            cell.setCellValue((Integer) valueOfCell);
        } else if (valueOfCell instanceof Long) {
            cell.setCellValue((Long) valueOfCell);
        } else if (valueOfCell instanceof String) {
            cell.setCellValue((String) valueOfCell);
        } else if(valueOfCell instanceof Double){
            cell.setCellValue((Double) valueOfCell);
        }
        cell.setCellStyle(style);
    }


    protected abstract void write();

    //metodo "main", llama al instancior de headers y lanza el metodo write
    //que es el encargado de recorrer lo q quiero exportar
    //y los coloca en las celdas
    //despues genere un publicador del archivo 'creo'
    //le escribe o postea el xlxs y cierra la edicion del xlxs
    //por el ultimo cierra el publicador
    public void generateExcelFile(HttpServletResponse response) throws IOException {
        instancedSheet();
        write();
        ServletOutputStream outputStream = response.getOutputStream();
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
    }

}