package com.bonartech.springbackend.Services.utils.reader.Reader.EjDeImplementaciones;

import com.bonartech.springbackend.Domain.Entity.Article;
import com.bonartech.springbackend.Services.DTOs.Article.response.ArticleExportDto;
import com.bonartech.springbackend.Services.utils.reader.ReportAbstract;
import com.bonartech.springbackend.enums.TypeExportArticle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.util.List;

public class ExportToExcel extends ReportAbstract {

    //Enum para switchear lo que quiero imprimir en las celdas
    private TypeExportArticle type;
    //Listas de objetos a exportar
    private List<ArticleExportDto> articles;

    public ExportToExcel(TypeExportArticle type, List<ArticleExportDto> articles) {
        super();
        this.type = type;
        this.articles = articles;
    }

    //se le puede enviar un estilo a la celda desde aca asi poder modificarla por eso el font
    //por cada elemento de mi lista creo una Row
    //por cada uno mando a imprimirse en la fila de las rows
    //le mando la row, el valor de la columna(que no deberia funcionar pero funciona) y el objeto
    protected void write() {
        int rowCount = 0;
        XSSFFont font = workbook.createFont();
        Integer columnCount = 0;
        for (ArticleExportDto article: articles) {
            rowCount++;
            Row row = sheet.createRow(rowCount);
            setCell(row, columnCount, article);
        }
    }

    //pinto las celdas con los valores del objeto
    //create cell esta defino en el padre(no modificar en lo posible)
    private void setCell(Row row, Integer column, ArticleExportDto article){
        createCell(row, column++, article.getCodeArticle(), null);
        createCell(row, column++, article.getName(), null);
        switch (this.type){
            case TypeExportArticle.providers:{
                if(article.getProvider() == null){
                    createCell(row, column++, "-", null);
                }else{
                    createCell(row, column++, article.getProvider(), null);
                }
                createCell(row, column++, article.getStock(), null);
                break;
            }
            case TypeExportArticle.budget:{
                createCell(row, column++, article.getPrice(), null);
                break;
            }
            default:
                break;
        }

    }

    //instancio una hoja
    //creo una columna
    //Genero un stylo tipo negrita para lso headers
    protected void instancedSheet(){
        this.sheet = workbook.createSheet();//atributo definido en el padre
        Row row = sheet.createRow(0);
        XSSFCellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setFontHeight(16);
        style.setFont(font);
        setHeaders(row, style);

    }

    //defino los headers de mi hoja y uso el enum para los distintos casos
    private void setHeaders(Row row, XSSFCellStyle style){
        createCell(row, 0, "Codigo de Articulo", style);
        createCell(row, 1, "Nombre", style);
        switch (this.type){
            case TypeExportArticle.providers:{
                createCell(row, 2, "Proveedor", style);
                createCell(row, 3, "Stock", style);
                break;
            }
            case TypeExportArticle.budget:{
                createCell(row, 2, "Precio", style);
            }
        }
    }
}