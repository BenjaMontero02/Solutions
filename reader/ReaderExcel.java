package com.bonartech.springbackend.Services.utils.reader;

import com.bonartech.springbackend.Domain.Entity.Article;
import com.bonartech.springbackend.exceptions.InvalidColumnExcel;
import com.bonartech.springbackend.exceptions.InvalidValueOfCell;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;


@NoArgsConstructor
public abstract class ReaderExcel<T> {

    protected Boolean isValid = true;
    protected List<T> entityToSave = new ArrayList<>();
    protected HashMap<String,String> columnHeaders = new HashMap<>();

    abstract void setHeadersHashMap();

    public abstract T getEntityToCreate();

    public List<T> readExcel(MultipartFile file) {

        try {
            //creo un archivo leible
            InputStream inputStreamFile = file.getInputStream();
            //obtengo el libro
            Workbook workBook = new XSSFWorkbook(inputStreamFile);
            //creo un contador para poder indicar en caso de error la fila
            int count = 1;
            //seteo los headers, cada clase q extiende tiene su propios headers para setear
            this.setHeadersHashMap();
            //recorro las columnas de mi excel
            for (Name name : workBook.getAllNames()) {
                //obtengo el nombre de la columna
                String columnName = name.getNameName();
                //obtengo la letra de la columna
                String column = letterOfColumnClear(name.getRefersToFormula());
                //si el nombre de la columna esta modificada, lanzo un error
                if(!columnHeaders.containsKey(columnName)){
                    throw new InvalidColumnExcel(column, columnName);
                }
                //si no agrego la letra a la que hace referencia esa columnn
                columnHeaders.put(columnName, column);
            }

            //obtengo la hoja
            Sheet sheet = workBook.getSheetAt(0);

            Iterator<Row> rowIterator = sheet.rowIterator();
            //recorro las filas siempre y cueando isValid sea true
            //isValid es falso cuando la fila anterior a la que estoy tiene todas sus celdas en blanco
            while(rowIterator.hasNext() && isValid) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIteratorsheet = row.cellIterator();
                T entity = getEntityToCreate();
                while(cellIteratorsheet.hasNext()) {
                    //obtengo la celda
                    Cell cell = cellIteratorsheet.next();
                    String cellAddress = cell.getAddress().formatAsString();
                    String columnLetter = cellAddress.replaceAll("[0-9]", ""); // Elimina los números de la dirección
                    //obtengo la key de mi hashMap para poder pasarsela a mi funcion setter de valores de la entidad
                    String keyName = this.getKeyName(columnLetter);
                    //obtengo el valor para poder setearlo a mi entidad
                    Object valueOfCell = getValueCell(cell);
                    this.setValueCellOnEntity(valueOfCell, keyName, entity, count, keyName);
                }
                this.entityToSave.add(entity);
                count++;
            }
            inputStreamFile.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return this.entityToSave;
    }

    private String getKeyName(String valorBuscado){
        String claveEncontrada = null;

        //obtengo el valor de la clave en base al valor del map
        for (String clave : columnHeaders.keySet()) {
            if (columnHeaders.get(clave).equals(valorBuscado)) {
                claveEncontrada = clave; // Obtener la clave correspondiente al valor
                break;
            }
        }

        return claveEncontrada;
    }

    //devuelvo un object con el valor de la celda
    private Object getValueCell(Cell cell){
        CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC:
                isValid = true;
                return cell.getNumericCellValue();
            case STRING:
                isValid = true;
                return cell.getStringCellValue();
            case BLANK:
                isValid = false;
                return null;
            default:
                return null;
        }
    }

    //verifico que que el valor que voy a gurdar en mi celda no sea true
    protected void verify(Object value, int row, String keyName){
        if(value == null){
            throw new InvalidValueOfCell(row, keyName);
        }
    }

    //seteo el valor de la celda a mi entidad
    abstract void setValueCellOnEntity(Object value, String type, T article, int row, String keyName);

    //elimino la letra de la columna
    protected String letterOfColumnClear(String refersToFormula){
        int indiceDolar = refersToFormula.indexOf('$');
        if (indiceDolar != -1 && indiceDolar < refersToFormula.length() - 1) {
            String letraColumna = String.valueOf(refersToFormula.charAt(indiceDolar + 1));
            return letraColumna;
        }else {
            return null;
        }
    }
}
