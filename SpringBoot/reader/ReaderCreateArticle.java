package com.bonartech.springbackend.Services.utils.reader.article;

import com.bonartech.springbackend.Domain.Entity.Article;
import com.bonartech.springbackend.Services.Article.ArticleService;
import com.bonartech.springbackend.Services.utils.reader.Reader;
import com.bonartech.springbackend.exceptions.InvalidColumnExcel;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.List;


/**
 * Class for create articles by .xlsx
 * @author linkedin.com/in/benjaminmontero/
 */
@NoArgsConstructor
public class ReaderCreateArticle extends Reader<Article> {

    /**
     * gets the entity
     * @return entity to work
     */
    public Article getEntityToCreate() {
        return new Article();
    }

    @Override
    public List<Article> readExcel(MultipartFile file) {

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
                Article entity = getEntityToCreate();
                isValid = false;
                while(cellIteratorsheet.hasNext()) {
                    //obtengo la celda
                    Cell cell = cellIteratorsheet.next();
                    String cellAddress = cell.getAddress().formatAsString();
                    String columnLetter = cellAddress.replaceAll("[0-9]", ""); // Elimina los números de la dirección
                    //obtengo la key de mi hashMap para poder pasarsela a mi funcion setter de valores de la entidad
                    String keyName = this.getKeyName(columnLetter);
                    //obtengo el valor para poder setearlo a mi entidad
                    Object valueOfCell = getValueCell(cell);
                    this.setValueCellOnEntity(valueOfCell, entity, count, keyName);
                    isValid = true;
                }
                if(isValid == false){break;}
                this.entityToSave.add(entity);
                count++;
            }
            inputStreamFile.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return this.entityToSave;
    }

    /**
     * defines the headers of the columns
     */
    protected void setHeadersHashMap(){
        columnHeaders.put("Titulo", null);
        columnHeaders.put("Descripcion", null);
        columnHeaders.put("PrecioUnitario", null);
        columnHeaders.put("Stock", null);
        columnHeaders.put("PrecioDeAdquisicion", null);
        columnHeaders.put("CodDeBarra", null);
    }

    /**
     *
     * @param value value to set to entity
     * @param article entity to be modified
     * @param row position of row
     * @param keyName column to reference
     */
    protected void setValueCellOnEntity(Object value, Article article, int row, String keyName){
        switch (keyName) {
            case "Titulo":
                this.verify(value, row, keyName);
                article.setTitle((String) value);
                break;
            case "Descripcion":
                if(value == null){
                    article.setDescription(null);
                }else{
                    article.setDescription((String) value);
                }
                break;
            case "PrecioUnitario":
                this.verify(value, row, keyName);
                article.setUnitPrice((Double) value);
                break;
            case "Stock":
                this.verify(value, row, keyName);
                double doubleValue = (Double) value;
                Long toIntegerValue = (long) doubleValue;
                article.setStockUnits(toIntegerValue);
                break;
            case "PrecioDeAdquisicion":
                this.verify(value, row, keyName);
                article.setAcquisitionCost((Double) value);
                break;
            case "CodDeBarra":
                if(value == null){
                    article.setCodBar(null);
                }else{
                    article.setCodBar((String) value);
                }
                break;
            default:
                break;
        }
    }
}