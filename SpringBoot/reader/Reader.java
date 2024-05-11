package com.bonartech.springbackend.Services.utils.reader;

import com.bonartech.springbackend.exceptions.InvalidValueOfCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.springframework.web.multipart.MultipartFile;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

public abstract class Reader<T> {
    /**
     * closes iteration of rows when they are invalid
     */
    protected Boolean isValid = true;

    /**
     * Entitys to return in the method readExcel
     */
    protected List<T> entityToSave = new ArrayList<>();
    /**
     * defines the column headings of the file sheet columns
     * key: Header, value: Column
     * Example: key:Title, value: B
     */
    protected HashMap<String,String> columnHeaders = new HashMap<>();

    public abstract List<T> readExcel(MultipartFile file);
    /**
     * Defines the keys for the attribute of class columnHeaders
     */
    protected abstract void setHeadersHashMap();
    /**
     * gets the key of the hash map
     * @param findValue value of hashMap to find
     * @return key found of hash map attribute
     */
    protected String getKeyName(String findValue){
        String keyFound = null;

        //obtengo el valor de la clave en base al valor del map
        for (String key : columnHeaders.keySet()) {
            if (columnHeaders.get(key).equals(findValue)) {
                keyFound = key; // Obtener la clave correspondiente al valor
                break;
            }
        }

        return keyFound;
    }

    /**
     * gets the value of cell
     * @param cell cell of row
     * @return value of the cell
     */
    //devuelvo un object con el valor de la celda
    protected Object getValueCell(Cell cell){
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

    /**
     * verify that the value of cell is not null
     * @param value cell to verify
     * @param row position of row
     * @param keyName column to reference
     * @throws InvalidValueOfCell if the value of cell is null
     */
    //verifico que que el valor que voy a gurdar en mi celda no sea null
    protected void verify(Object value, int row, String keyName){
        if(value == null){
            throw new InvalidValueOfCell(row, keyName);
        }
    }

    /**
     * set in the entity the value
     * @param value value to set to entity
     * @param entity entity to be modified
     * @param row position of row
     * @param keyName column to reference
     */
    //seteo el valor de la celda a mi entidad
    protected abstract void setValueCellOnEntity(Object value, T entity, int row, String keyName);

    /**
     * gets the letter of the formula of the column of sheet
     * Example: Sheet1$B:$B, return B
     * @param refersToFormula formula of the column of sheet
     * @return the letter of the formula
     */
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