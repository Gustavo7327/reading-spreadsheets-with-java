package com.planilhas.tratarplanilhas;

import java.io.File;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.util.List;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;

import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.google.gson.Gson;

public class RequestLib {
    public static void main(String[] args) {
        try {

            Workbook wb = new Workbook("/home/guilherme/Documentos/ProjetoBibliotecaUtils/BibliotecaPlanilha.xlsx");
            wb.getWorksheets().add("DadosRequisitados");
            Worksheet antiga = wb.getWorksheets().get("DadosNaoRepetidosENecessarios");
            Worksheet nova = wb.getWorksheets().get("DadosRequisitados");
            
            nova.getCells().get(0, 0).setValue("TITULO");
            nova.getCells().get(0, 1).setValue("AUTOR");
            nova.getCells().get(0, 2).setValue("ANO");
            nova.getCells().get(0, 3).setValue("IMAGEM");
            nova.getCells().get(0, 4).setValue("AMOSTRA");
            nova.getCells().get(0, 5).setValue("SINOPSE");
            nova.getCells().get(0, 6).setValue("GENERO");
            nova.getCells().get(0, 7).setValue("AVALIACAO");
            nova.getCells().get(0, 8).setValue("NUMPAGINAS");

            for(int row = 1; row < antiga.getCells().getMaxDataRow(); row++){
                String urlAPI = "https://www.googleapis.com/books/v1/volumes?q=" + antiga.getCells().get(row, 0).getStringValue().replace(" ", "");

                HttpClient client = HttpClient.newHttpClient();
                HttpRequest request = HttpRequest.newBuilder().uri(URI.create(urlAPI)).build();
                HttpResponse<String> response = client.send(request, HttpResponse.BodyHandlers.ofString());

                Gson gson = new Gson();
                Volume volume = gson.fromJson(response.body(), Volume.class);

                if(volume != null && volume.getItems() != null && !volume.getItems().isEmpty()){
                    Item item = volume.getItems().get(0);

                    nova.getCells().get(row, 0).setValue(antiga.getCells().get(row, 0).getStringValue());
                    if(item.getVolumeInfo().getAuthors() != null){
                        nova.getCells().get(row, 1).setValue(item.getVolumeInfo().getAuthors());
                    }

                    if(item.getVolumeInfo().getPublishedDate() != null){
                        nova.getCells().get(row, 2).setValue(item.getVolumeInfo().getPublishedDate());
                    }

                    if(item.getVolumeInfo().getImageLinks() != null){
                        nova.getCells().get(row, 3).setValue(item.getVolumeInfo().getImageLinks().getThumbnail());
                    }

                    if(item.getAccessInfo().getWebReaderLink() != null){
                        nova.getCells().get(row, 4).setValue(item.getAccessInfo().getWebReaderLink());
                    }

                    if(item.getVolumeInfo().getDescription() != null){
                        nova.getCells().get(row, 5).setValue(item.getVolumeInfo().getDescription());
                    }

                    if(item.getVolumeInfo().getCategories() != null){
                        nova.getCells().get(row, 6).setValue(item.getVolumeInfo().getCategories());
                    }

                    if(item.getVolumeInfo().getAverageRating() != null){
                        nova.getCells().get(row, 7).setValue(item.getVolumeInfo().getAverageRating());
                    }

                    if(item.getVolumeInfo().getPageCount() != null){
                        nova.getCells().get(row, 8).setValue(item.getVolumeInfo().getPageCount());
                    }
                    
                } 
            }
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setDialogTitle("Save Workbook");
            FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel Files (*.xlsx)", "xlsx");
            fileChooser.setFileFilter(filter);

            int result = fileChooser.showSaveDialog(null);
            if(result == JFileChooser.APPROVE_OPTION){
                File file = fileChooser.getSelectedFile();
                wb.save(file.getAbsolutePath(), SaveFormat.XLSX);
                System.out.println("Workbook save in " + file.getAbsolutePath());
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    private class Volume{
        private List<Item> items;

        public List<Item> getItems(){
            return items;
        }
    }


    private class Item{
        private VolumeInfo volumeInfo;
        private AccessInfo accessInfo;

        public VolumeInfo getVolumeInfo() {
            return volumeInfo;
        }
        public AccessInfo getAccessInfo() {
            return accessInfo;
        }
    }

    private class VolumeInfo{
        private ImageLinks imageLinks;
        private String description;
        private List<String> authors;
        private String publishedDate;
        private List<String> categories;
        private Double averageRating;
        private Integer pageCount;

        public ImageLinks getImageLinks() {
            return imageLinks;
        }
        public String getDescription() {
            return description;
        }
        public List<String> getAuthors() {
            return authors;
        }
        public String getPublishedDate() {
            return publishedDate;
        }
        public List<String> getCategories() {
            return categories;
        }
        public Double getAverageRating() {
            return averageRating;
        }
        public Integer getPageCount() {
            return pageCount;
        }
    }

    private class ImageLinks{
        private String thumbnail;

        public String getThumbnail() {
            return thumbnail;
        }
    }

    private class AccessInfo{
        private String webReaderLink;

        public String getWebReaderLink() {
            return webReaderLink;
        }
    }
}
