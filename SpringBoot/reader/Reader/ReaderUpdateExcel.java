package com.bonartech.springbackend.Services.utils.reader;

import com.bonartech.springbackend.Domain.Entity.Article;
import com.bonartech.springbackend.exceptions.InvalidColumnExcel;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public abstract class ReaderUpdateExcel<T> extends Reader<T>{

    protected List<T> listOfEntitysValidated = new ArrayList<T>();

    @Override
    protected void setHeadersHashMap(){};

    protected abstract T getEntityById(Object id);

    @Override
    protected void setValueCellOnEntity(Object value, T entity, int row, String keyName) {

    }

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
            int countCell = 0;
            //obtengo la hoja
            Sheet sheet = workBook.getSheetAt(0);

            Iterator<Row> rowIterator = sheet.rowIterator();
            //recorro las filas siempre y cueando isValid sea true
            //isValid es falso cuando la fila anterior a la que estoy tiene todas sus celdas en blanco
            while(rowIterator.hasNext() && isValid) {
                Row row = rowIterator.next();
                //lo utilizo para recorrer las celdas del identificador
                Iterator<Cell> cellIterator = row.cellIterator();
                countCell = 0;
                while (cellIterator.hasNext()) {
                    //obtengo la celda
                    Cell cell = cellIterator.next();
                    String cellAddress = cell.getAddress().formatAsString();
                    String columnLetter = cellAddress.replaceAll("[0-9]", ""); // Elimina los números de la dirección
                    //obtengo la key de mi hashMap para poder pasarsela a mi funcion setter de valores de la entidad
                    String keyName = this.getKeyName(columnLetter);
                    //obtengo el valor para poder setearlo a mi entidad
                    if (countCell == 0) {
                        Object valueOfCell = getValueCell(cell);
                        T entity = getEntityById(valueOfCell);
                        entityToSave.add(entity);
                    }else{
                        Object valueOfCell = getValueCell(cell);
                        setValueCellOnEntity(valueOfCell, entityToSave.getLast(), count, keyName);
                    }

                    countCell++;
                }
            }


            inputStreamFile.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return this.entityToSave;
    }


}