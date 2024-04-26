package com.bonartech.springbackend.Services.utils.reader;

import com.bonartech.springbackend.Domain.Entity.Article;
import lombok.NoArgsConstructor;

@NoArgsConstructor
public class ReaderCreateArticle extends ReaderExcel<Article>{

    @Override
    public Article getEntityToCreate() {
        return new Article();
    }

    protected void setHeadersHashMap(){
        columnHeaders.put("Titulo", null);
        columnHeaders.put("Descripcion", null);
        columnHeaders.put("PrecioUnitario", null);
        columnHeaders.put("Stock", null);
        columnHeaders.put("PrecioDeAdquisicion", null);
    }

    protected void setValueCellOnEntity(Object value, String type, Article article, int row, String keyName){
        switch (type) {
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
                int toIntegerValue = (int) doubleValue;
                article.setStockUnits(toIntegerValue);
                break;
            case "PrecioDeAdquisicion":
                this.verify(value, row, keyName);
                article.setAcquisitionCost((Double) value);
                break;
            default:
                break;
        }
    }
}
