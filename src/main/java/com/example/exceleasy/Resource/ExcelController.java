package com.example.exceleasy.Resource;

import com.example.exceleasy.Model.Model;
import com.example.exceleasy.Model.SubModel;
import com.example.exceleasy.Repository.ExcelRepo;
import com.example.exceleasy.Repository.SubRepo;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Configuration;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.data.domain.Example;
import org.springframework.data.mongodb.core.convert.MappingMongoConverter;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;

@SuppressWarnings({"unused","rawtypes"})
@Configuration
@RestController
public class ExcelController {

    String filename = "";

    @Autowired
    void setMapKeyDotReplacement(MappingMongoConverter mappingMongoConverter) {
        mappingMongoConverter.setMapKeyDotReplacement("_");
    }

    @Autowired
    private ExcelRepo repo;

    @Autowired
    private SubRepo subRepo;

    //operations on base file

    @GetMapping("/rules")
    public ArrayList<String> rulebook() {
        ArrayList<String> arrayList = new ArrayList<>();

        arrayList.add("1) All the files should be in .xlsx extension");
        arrayList.add("2) Text fields should only contain text values and Numeral fields should only contain numerical values." +
                " E.g. Description of items/ Units should not contain even a single numerical value." +
                " Rate/ Amount/ Quantity should never contain a text value. If this is not ensured then the software may crash. ");
        arrayList.add("3) The name of fields should exactly be as follows: 'SOR', 'DESCRIPTION OF ITEMS', 'Quantity', 'Units', 'Rate', 'Amount' ");
        arrayList.add("4) In base file the sheet from which data is to be extracted should always be the first sheet");
        arrayList.add("5) The name of the sheets in base file/ sub files should exactly be: 'Civil', 'Electrical', 'Sanitary' ");
        arrayList.add("6) The blue coloured rows depict the rows in the sub file which are not in the base file");

        return arrayList;
    }

    @GetMapping("/commands")
    public ArrayList<String> commandBook() {
        ArrayList<String> arrayList = new ArrayList<>();

        arrayList.add(" Base file commands: ");
        arrayList.add("1) http://3.143.215.20:8080/addbasefile" + " : Add base file (attach the file) ");
        arrayList.add("2) http://3.143.215.20:8080/finditem/item_id_here" + " : Get the details of the item by it's id");
        arrayList.add("3) http://3.143.215.20:8080/updateitem/item_id_here" + " : Update the details of the item by it's id");
        arrayList.add("5) http://3.143.215.20:8080/getbasefile" + " : Get the contents of the base file");
        arrayList.add("6) http://3.143.215.20:8080/deleteitem/item_id_here" + " : Delete an item by it's id");
        arrayList.add("7) http://3.143.215.20:8080/deleteallitems" + " : Delete all the items in the base file");
        arrayList.add("");
        arrayList.add(" Sub file commands: ");
        arrayList.add("8) http://3.143.215.20:8080/addfile" + " : Add a sub file (attach the file)");
        arrayList.add("9) http://3.143.215.20:8080/getsubfile" + " : Get the contents of the sub file");
        arrayList.add("10) http://3.143.215.20:8080/clearsubfile" + " : Delete all the items in the sub file");
        arrayList.add("11) http://3.143.215.20:8080/reset" + " : Set the Quantity/ Amount of all the items in the sub file to zero");
        arrayList.add("12) http://3.143.215.20:8080/getResultFile" + " : Get the final result file");

        return arrayList;
    }

    @PostMapping("/addBaseFile")
    public void saveBase(@RequestParam("file") MultipartFile file) {
        try {

            Path tempDir = Files.createTempDirectory("");

            File tempFile = tempDir.resolve(Objects.requireNonNull(file.getOriginalFilename())).toFile();

            file.transferTo(tempFile);

            
            XSSFWorkbook wb = new XSSFWorkbook(tempFile);

            XSSFSheet sheet = wb.getSheetAt(0);

            String SOR = "", DESCRIPTION_OF_ITEMS = "", Unit = "", Rate = "";

            Row row2 = sheet.getRow(1); //default

            int x = 0;

            for (Row row : sheet) {
                Iterator<Cell> cellIterator2 = row.cellIterator();   //iterating over each column
                while (cellIterator2.hasNext()) {
                    Cell cell = cellIterator2.next();

                    if (cell.getStringCellValue().equals("SOR")) {
                        row2 = row;
                        x = 1;
                        break;
                    }
                }
                if (x == 1) {
                    break;
                }
            }

            Iterator<Cell> cellIterator = row2.cellIterator();   //iterating over each column
            while (cellIterator.hasNext()) {

                Cell cell = cellIterator.next();

                if (cell.getStringCellValue().trim().equals("SOR")) {
                    SOR = CellReference.convertNumToColString(cell.getColumnIndex()); //A
                } else if (cell.getStringCellValue().trim().equals("Description Of Items")) //B
                {
                    DESCRIPTION_OF_ITEMS = CellReference.convertNumToColString(cell.getColumnIndex());
                } else if (cell.getStringCellValue().trim().equals("Unit")) //C
                {
                    Unit = CellReference.convertNumToColString(cell.getColumnIndex());
                } else if (cell.getStringCellValue().trim().equals("Rate")) //F
                {
                    Rate = CellReference.convertNumToColString(cell.getColumnIndex());
                }
            }

            DataFormatter formatter = new DataFormatter(Locale.US);

            for (Row row : sheet) {

                if (row.getRowNum() > row2.getRowNum()) {
                    Model model = new Model();

                    Iterator<Cell> cellIterator2 = row.cellIterator();   //iterating over each column
                    while (cellIterator2.hasNext()) {
                        Cell cell = cellIterator2.next();

                        if (CellReference.convertNumToColString(cell.getColumnIndex()).equals(SOR)) {

                            FormulaEvaluator formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();

                            if (formulaEvaluator.evaluateInCell(cell).getCellTypeEnum().equals(CellType.NUMERIC)) {
                                model.setSOR(formatter.formatCellValue(cell));
                            } else {
                                model.setSOR(cell.getStringCellValue());
                            }

                        } else if (CellReference.convertNumToColString(cell.getColumnIndex()).equals(DESCRIPTION_OF_ITEMS)) {

                            model.setDESCRIPTION_OF_ITEMS(cell.getStringCellValue());
                        } else if (CellReference.convertNumToColString(cell.getColumnIndex()).equals(Unit)) {

                            model.setUnit(cell.getStringCellValue());
                        } else if (CellReference.convertNumToColString(cell.getColumnIndex()).equals(Rate)) {

                            FormulaEvaluator formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();

                            if (formulaEvaluator.evaluateInCell(cell).getCellTypeEnum().equals(CellType.NUMERIC)) {
                                model.setRate(cell.getNumericCellValue());
                            } else {
                                model.setRate(0.0);
                            }
                        }
                    }
                    if (!repo.exists(Example.of(model)) && !model.getSOR().matches(".*[a-zA-Z]+.*") &&
                            !model.getDESCRIPTION_OF_ITEMS().trim().isEmpty() && !model.getSOR().trim().isEmpty()) {

                        repo.save(model);
                        //  System.out.println(model.getSOR());
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @GetMapping("/findItem/{id}")
    public Optional<Model> getItem(@PathVariable String id) {
        return repo.findById(id);
    }

    @PostMapping("/updateItem/{id}")
    public String updateItem(@RequestBody Model model) {
        repo.save(model);
        return "Updated Successfully";
    }

    @PostMapping("/updateRate")
    public String updateRate(@RequestBody Model model1) {
        try {
            if (repo.findById(model1.getSOR()).isPresent()) {
                Model model = repo.findById(model1.getSOR()).get();
                model.setRate(model1.getRate());
                repo.save(model);
                return "Updated Successfully";
            }
        } catch (Exception e) {
            e.printStackTrace();
            return "No item exists by that ID";
        }
        return "";
    }

    @GetMapping("/getBaseFile")
    public List<Model> getBaseFile() {
        return repo.findAll();
    }

    @DeleteMapping("/deleteItem/{id}")
    public String deleteItem(@PathVariable String id) {
        repo.deleteById(id);
        return "Item Deleted Successfully!";
    }

    @DeleteMapping("/deleteAllItems")
    public String deleteAll() {
        repo.deleteAll();
        return "All items were Deleted Successfully!";
    }

    //operations on sub-files

    @GetMapping("/getSubFile")
    public List<SubModel> getSubFile() {
        return subRepo.findAll();
    }

    @DeleteMapping("/clearSubFile")
    public String deleteAllFromSubFile() {
        subRepo.deleteAll();
        return "All items were Deleted Successfully!";
    }

    @PostMapping("/reset")
    public String reset() {

        List<SubModel> list = getSubFile();

        for (SubModel model : list) {
            model.setQuantity(0.0);
            model.setAmount(0.0);
            subRepo.save(model);
        }
        return "Reset Successfully!";
    }

    @PostMapping("/addFile")
    public void addFile(@RequestParam("file") MultipartFile file) {

        try {
            Path tempDir = Files.createTempDirectory("");
            File tempFile = tempDir.resolve(Objects.requireNonNull(file.getOriginalFilename())).toFile();
            file.transferTo(tempFile);

            filename = file.getOriginalFilename();

            XSSFWorkbook wb = new XSSFWorkbook(tempFile);

            civil(wb);
            electrical(wb);
            sanitary(wb);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void civil(XSSFWorkbook wb) {
        try {
            XSSFSheet sheet = wb.getSheet("Civil");

            if (sheet != null) {

                String SOR_C = "", Quantity_C = "", Amount_C = "", Desc_C = "";
                FormulaEvaluator formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();
                DataFormatter formatter = new DataFormatter(Locale.US);

                Row row2 = sheet.getRow(1); //default

                String temp;
                int x = 0;

                for (Row row : sheet) {
                    Iterator<Cell> cellIterator2 = row.cellIterator();   //iterating over each column
                    while (cellIterator2.hasNext()) {
                        Cell cell = cellIterator2.next();

                        if (formulaEvaluator.evaluateInCell(cell).getCellTypeEnum().equals(CellType.NUMERIC)) {
                            temp = formatter.formatCellValue(cell);
                        } else {
                            temp = cell.getStringCellValue();
                        }
                        if (temp.equals("SOR")) {
                            row2 = row;
                            x = 1;
                            break;
                        }
                    }
                    if (x == 1) {
                        break;
                    }
                }

                Iterator<Cell> cellIterator1 = row2.cellIterator();   //iterating over each column
                while (cellIterator1.hasNext()) {

                    Cell cell = cellIterator1.next();

                    if (cell.getStringCellValue().trim().equals("SOR")) {
                        SOR_C = CellReference.convertNumToColString(cell.getColumnIndex()); //B

                    } else if (cell.getStringCellValue().trim().equals("Description Of Items")) {
                        Desc_C = CellReference.convertNumToColString(cell.getColumnIndex()); //C

                    } else if (cell.getStringCellValue().trim().equals("Quantity")) {
                        Quantity_C = CellReference.convertNumToColString(cell.getColumnIndex()); //H

                    } else if (cell.getStringCellValue().trim().equals("Amount")) {
                        Amount_C = CellReference.convertNumToColString(cell.getColumnIndex()); //K

                    }
                }

                String SOR = "";
                double Quantity = -1.0;
                double amount = -1.0;
                String Desc = "";

                for (Row row : sheet) {

                    if (row.getRowNum() > row2.getRowNum()) {

                        Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column
                        while (cellIterator.hasNext()) {
                            Cell cell = cellIterator.next();

                            if (CellReference.convertNumToColString(cell.getColumnIndex()).equals(SOR_C)) {

                                if (formulaEvaluator.evaluateInCell(cell).getCellTypeEnum().equals(CellType.NUMERIC)) {
                                    if (!formatter.formatCellValue(cell).isEmpty()) {
                                        SOR = formatter.formatCellValue(cell);
                                    }
                                } else {
                                    if (!cell.getStringCellValue().isEmpty()) {
                                        SOR = cell.getStringCellValue();
                                    }
                                }

                            } else if (CellReference.convertNumToColString(cell.getColumnIndex()).equals(Desc_C)) {

                                if (formulaEvaluator.evaluateInCell(cell).getCellTypeEnum().equals(CellType.NUMERIC)) {
                                    if (!formatter.formatCellValue(cell).isEmpty()) {
                                        Desc = formatter.formatCellValue(cell);
                                    }
                                } else {
                                    if (!cell.getStringCellValue().isEmpty()) {
                                        Desc = cell.getStringCellValue();
                                    }
                                }

                            } else if (CellReference.convertNumToColString(cell.getColumnIndex()).equals(Quantity_C)) {

                                double temp_quantity = cell.getNumericCellValue();

                                Quantity = BigDecimal.valueOf(temp_quantity)
                                        .setScale(3, RoundingMode.HALF_UP)
                                        .doubleValue();

                                //   System.out.println(Quantity);
                            } else if (CellReference.convertNumToColString(cell.getColumnIndex()).equals(Amount_C)) {

                                double temp_amount = cell.getNumericCellValue();

                                amount = BigDecimal.valueOf(temp_amount)
                                        .setScale(3, RoundingMode.HALF_UP)
                                        .doubleValue();

                                //System.out.println(amount);
                            }

                        }

                        if (!SOR.isEmpty() && amount != 0.0) {

                            Double rate = 0.0;
                            boolean status;

                            //    System.out.println(SOR + " "+ Quantity);

                            if (searchRate(SOR).isPresent()) {
                                Model model = searchRate(SOR).get();
                                rate = model.getRate();
                                status = true; //item exists in main db
                            } else {
                                status = false; //item doesn't exist in main db
                            }
                            SubModel subModel;

                            HashMap<String, Double> hash = new HashMap<>();

                            if (searchQuantity(SOR).isPresent()) {

                                subModel = searchQuantity(SOR).get();
                                double quantity = subModel.getQuantity();

                                double new_Quantity = quantity + Quantity;
                                double temp_Amount = new_Quantity * rate;

                                double new_Amount = BigDecimal.valueOf(temp_Amount)
                                        .setScale(3, RoundingMode.HALF_UP)
                                        .doubleValue();

                                hash = subModel.getConstituentSheets();

                                if (hash.containsKey(filename)) {
                                    double temp_Quantity = hash.get(filename);
                                    hash.put(filename, temp_Quantity + Quantity);
                                } else {
                                    hash.put(filename, Quantity);
                                }

                                subModel.setConstituentSheets(hash);
                                subModel.setQuantity(new_Quantity);
                                subModel.setAmount(new_Amount);

                            } else {

                                double temp_Amount = Quantity * rate;
                                subModel = new SubModel();

                                double Amount = BigDecimal.valueOf(temp_Amount)
                                        .setScale(3, RoundingMode.HALF_UP)
                                        .doubleValue();

                                hash.put(filename, Quantity);

                                subModel.setConstituentSheets(hash);
                                subModel.setSOR(SOR);
                                subModel.setQuantity(Quantity);
                                subModel.setAmount(Amount);
                            }

                            if (!status) {
                                subModel.setDESCRIPTION_OF_ITEMS(Desc);
                            } else {
                                subModel.setDESCRIPTION_OF_ITEMS("");
                            }

                            subModel.setStatus(status);
                            subRepo.save(subModel);

                            SOR = "";
                            Quantity = -1.0;
                        }
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void electrical(XSSFWorkbook wb) {
        try {
            XSSFSheet sheet = wb.getSheet("Electrical");


            if (sheet != null) {

                String SOR_C = "", Quantity_C = "", Amount_C = "", Desc_C = "";
                FormulaEvaluator formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();
                DataFormatter formatter = new DataFormatter(Locale.US);

                Row row2 = sheet.getRow(1); //default

                String temp;
                int x = 0;

                for (Row row : sheet) {
                    Iterator<Cell> cellIterator2 = row.cellIterator();   //iterating over each column
                    while (cellIterator2.hasNext()) {
                        Cell cell = cellIterator2.next();

                        if (formulaEvaluator.evaluateInCell(cell).getCellTypeEnum().equals(CellType.NUMERIC)) {
                            temp = formatter.formatCellValue(cell);
                        } else {
                            temp = cell.getStringCellValue();
                        }
                        if (temp.equals("SOR")) {
                            row2 = row;
                            x = 1;
                            break;
                        }
                    }
                    if (x == 1) {
                        break;
                    }
                }


                Iterator<Cell> cellIterator1 = row2.cellIterator();   //iterating over each column
                while (cellIterator1.hasNext()) {

                    Cell cell = cellIterator1.next();

                    if (cell.getStringCellValue().trim().equals("SOR")) {
                        SOR_C = CellReference.convertNumToColString(cell.getColumnIndex()); //B

                    } else if (cell.getStringCellValue().trim().equals("Description Of Items")) {
                        Desc_C = CellReference.convertNumToColString(cell.getColumnIndex()); //C

                    } else if (cell.getStringCellValue().trim().equals("Quantity")) {
                        Quantity_C = CellReference.convertNumToColString(cell.getColumnIndex()); //H

                    } else if (cell.getStringCellValue().trim().equals("Amount")) {
                        Amount_C = CellReference.convertNumToColString(cell.getColumnIndex()); //K

                    }
                }

                String SOR = "";
                double Quantity = -1.0;
                double amount = -1.0;
                String Desc = "";

                for (Row row : sheet) {

                    if (row.getRowNum() > row2.getRowNum()) {

                        Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column
                        while (cellIterator.hasNext()) {
                            Cell cell = cellIterator.next();

                            if (CellReference.convertNumToColString(cell.getColumnIndex()).equals(SOR_C)) {

                                if (formulaEvaluator.evaluateInCell(cell).getCellTypeEnum().equals(CellType.NUMERIC)) {
                                    if (!formatter.formatCellValue(cell).isEmpty()) {
                                        SOR = formatter.formatCellValue(cell);
                                    }
                                } else {
                                    if (!cell.getStringCellValue().isEmpty()) {
                                        SOR = cell.getStringCellValue();
                                    }
                                }

                            } else if (CellReference.convertNumToColString(cell.getColumnIndex()).equals(Desc_C)) {

                                if (formulaEvaluator.evaluateInCell(cell).getCellTypeEnum().equals(CellType.NUMERIC)) {
                                    if (!formatter.formatCellValue(cell).isEmpty()) {
                                        Desc = formatter.formatCellValue(cell);
                                    }
                                } else {
                                    if (!cell.getStringCellValue().isEmpty()) {
                                        Desc = cell.getStringCellValue();
                                    }
                                }

                            } else if (CellReference.convertNumToColString(cell.getColumnIndex()).equals(Quantity_C)) {

                                double temp_quantity = cell.getNumericCellValue();

                                Quantity = BigDecimal.valueOf(temp_quantity)
                                        .setScale(3, RoundingMode.HALF_UP)
                                        .doubleValue();

                                //   System.out.println(Quantity);
                            } else if (CellReference.convertNumToColString(cell.getColumnIndex()).equals(Amount_C)) {

                                double temp_amount = cell.getNumericCellValue();

                                amount = BigDecimal.valueOf(temp_amount)
                                        .setScale(3, RoundingMode.HALF_UP)
                                        .doubleValue();

                                //System.out.println(amount);
                            }

                        }

                        if (!SOR.isEmpty() && amount != 0.0) {

                            Double rate = 0.0;
                            boolean status;

                            //    System.out.println(SOR + " "+ Quantity);

                            if (searchRate(SOR).isPresent()) {
                                Model model = searchRate(SOR).get();
                                rate = model.getRate();
                                status = true; //item exists in main db
                            } else {
                                status = false; //item doesn't exist in main db
                            }
                            SubModel subModel;

                            HashMap<String, Double> hash = new HashMap<>();

                            if (searchQuantity(SOR).isPresent()) {

                                subModel = searchQuantity(SOR).get();
                                double quantity = subModel.getQuantity();

                                double new_Quantity = quantity + Quantity;
                                double temp_Amount = new_Quantity * rate;

                                double new_Amount = BigDecimal.valueOf(temp_Amount)
                                        .setScale(3, RoundingMode.HALF_UP)
                                        .doubleValue();

                                hash = subModel.getConstituentSheets();

                                if (hash.containsKey(filename)) {
                                    double temp_Quantity = hash.get(filename);
                                    hash.put(filename, temp_Quantity + Quantity);
                                } else {
                                    hash.put(filename, Quantity);
                                }

                                subModel.setConstituentSheets(hash);
                                subModel.setQuantity(new_Quantity);
                                subModel.setAmount(new_Amount);

                            } else {

                                double temp_Amount = Quantity * rate;
                                subModel = new SubModel();

                                double Amount = BigDecimal.valueOf(temp_Amount)
                                        .setScale(3, RoundingMode.HALF_UP)
                                        .doubleValue();

                                hash.put(filename, Quantity);

                                subModel.setConstituentSheets(hash);
                                subModel.setSOR(SOR);
                                subModel.setQuantity(Quantity);
                                subModel.setAmount(Amount);
                            }

                            if (!status) {
                                subModel.setDESCRIPTION_OF_ITEMS(Desc);
                            } else {
                                subModel.setDESCRIPTION_OF_ITEMS("");
                            }

                            subModel.setStatus(status);
                            subRepo.save(subModel);

                            SOR = "";
                            Quantity = -1.0;
                        }
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void sanitary(XSSFWorkbook wb) {
        try {
            XSSFSheet sheet = wb.getSheet("Sanitary");


            if (sheet != null) {

                String SOR_C = "", Quantity_C = "", Amount_C = "", Desc_C = "";
                FormulaEvaluator formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();
                DataFormatter formatter = new DataFormatter(Locale.US);

                Row row2 = sheet.getRow(1); //default

                String temp;
                int x = 0;

                for (Row row : sheet) {
                    Iterator<Cell> cellIterator2 = row.cellIterator();   //iterating over each column
                    while (cellIterator2.hasNext()) {
                        Cell cell = cellIterator2.next();

                        if (formulaEvaluator.evaluateInCell(cell).getCellTypeEnum().equals(CellType.NUMERIC)) {
                            temp = formatter.formatCellValue(cell);
                        } else {
                            temp = cell.getStringCellValue();
                        }
                        if (temp.equals("SOR")) {
                            row2 = row;
                            x = 1;
                            break;
                        }
                    }
                    if (x == 1) {
                        break;
                    }
                }


                Iterator<Cell> cellIterator1 = row2.cellIterator();   //iterating over each column
                while (cellIterator1.hasNext()) {

                    Cell cell = cellIterator1.next();

                    if (cell.getStringCellValue().trim().equals("SOR")) {
                        SOR_C = CellReference.convertNumToColString(cell.getColumnIndex()); //B

                    } else if (cell.getStringCellValue().trim().equals("Description Of Items")) {
                        Desc_C = CellReference.convertNumToColString(cell.getColumnIndex()); //C

                    } else if (cell.getStringCellValue().trim().equals("Quantity")) {
                        Quantity_C = CellReference.convertNumToColString(cell.getColumnIndex()); //H

                    } else if (cell.getStringCellValue().trim().equals("Amount")) {
                        Amount_C = CellReference.convertNumToColString(cell.getColumnIndex()); //K

                    }
                }

                String SOR = "";
                double Quantity = -1.0;
                double amount = -1.0;
                String Desc = "";

                for (Row row : sheet) {

                    if (row.getRowNum() > row2.getRowNum()) {

                        Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column
                        while (cellIterator.hasNext()) {
                            Cell cell = cellIterator.next();

                            if (CellReference.convertNumToColString(cell.getColumnIndex()).equals(SOR_C)) {

                                if (formulaEvaluator.evaluateInCell(cell).getCellTypeEnum().equals(CellType.NUMERIC)) {
                                    if (!formatter.formatCellValue(cell).isEmpty()) {
                                        SOR = formatter.formatCellValue(cell);
                                    }
                                } else {
                                    if (!cell.getStringCellValue().isEmpty()) {
                                        SOR = cell.getStringCellValue();
                                    }
                                }

                            } else if (CellReference.convertNumToColString(cell.getColumnIndex()).equals(Desc_C)) {

                                if (formulaEvaluator.evaluateInCell(cell).getCellTypeEnum().equals(CellType.NUMERIC)) {
                                    if (!formatter.formatCellValue(cell).isEmpty()) {
                                        Desc = formatter.formatCellValue(cell);
                                    }
                                } else {
                                    if (!cell.getStringCellValue().isEmpty()) {
                                        Desc = cell.getStringCellValue();
                                    }
                                }

                            } else if (CellReference.convertNumToColString(cell.getColumnIndex()).equals(Quantity_C)) {

                                double temp_quantity = cell.getNumericCellValue();

                                Quantity = BigDecimal.valueOf(temp_quantity)
                                        .setScale(3, RoundingMode.HALF_UP)
                                        .doubleValue();

                                //   System.out.println(Quantity);
                            } else if (CellReference.convertNumToColString(cell.getColumnIndex()).equals(Amount_C)) {

                                double temp_amount = cell.getNumericCellValue();

                                amount = BigDecimal.valueOf(temp_amount)
                                        .setScale(3, RoundingMode.HALF_UP)
                                        .doubleValue();

                                //System.out.println(amount);
                            }

                        }

                        if (!SOR.isEmpty() && amount != 0.0) {

                            Double rate = 0.0;
                            boolean status;

                            //    System.out.println(SOR + " "+ Quantity);

                            if (searchRate(SOR).isPresent()) {
                                Model model = searchRate(SOR).get();
                                rate = model.getRate();
                                status = true; //item exists in main db
                            } else {
                                status = false; //item doesn't exist in main db
                            }
                            SubModel subModel;

                            HashMap<String, Double> hash = new HashMap<>();

                            if (searchQuantity(SOR).isPresent()) {

                                subModel = searchQuantity(SOR).get();
                                double quantity = subModel.getQuantity();

                                double new_Quantity = quantity + Quantity;
                                double temp_Amount = new_Quantity * rate;

                                double new_Amount = BigDecimal.valueOf(temp_Amount)
                                        .setScale(3, RoundingMode.HALF_UP)
                                        .doubleValue();

                                hash = subModel.getConstituentSheets();

                                if (hash.containsKey(filename)) {
                                    double temp_Quantity = hash.get(filename);
                                    hash.put(filename, temp_Quantity + Quantity);
                                } else {
                                    hash.put(filename, Quantity);
                                }

                                subModel.setConstituentSheets(hash);
                                subModel.setQuantity(new_Quantity);
                                subModel.setAmount(new_Amount);

                            } else {

                                double temp_Amount = Quantity * rate;
                                subModel = new SubModel();

                                double Amount = BigDecimal.valueOf(temp_Amount)
                                        .setScale(3, RoundingMode.HALF_UP)
                                        .doubleValue();

                                hash.put(filename, Quantity);

                                subModel.setConstituentSheets(hash);
                                subModel.setSOR(SOR);
                                subModel.setQuantity(Quantity);
                                subModel.setAmount(Amount);
                            }

                            if (!status) {
                                subModel.setDESCRIPTION_OF_ITEMS(Desc);
                            } else {
                                subModel.setDESCRIPTION_OF_ITEMS("");
                            }

                            subModel.setStatus(status);
                            subRepo.save(subModel);

                            SOR = "";
                            Quantity = -1.0;
                        }
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @GetMapping("/searchRate") // not for direct use
    public Optional<Model> searchRate(@RequestBody String SOR) {
        return repo.findById(SOR);
    }

    @GetMapping("/searchQuantity") // not for direct use
    public Optional<SubModel> searchQuantity(@RequestBody String SOR) {
        return subRepo.findById(SOR);
    }

    @GetMapping("/getResultFile")
    public ResponseEntity<ByteArrayResource> getRF(HttpServletResponse response) {
        try {

            ByteArrayOutputStream stream = new ByteArrayOutputStream();
            XSSFWorkbook workbook = new XSSFWorkbook();

            XSSFSheet sheet = workbook.createSheet("Final");

            XSSFRow rowHead = sheet.createRow((short) 0);

            rowHead.createCell(1).setCellValue("SOR");
            rowHead.createCell(2).setCellValue("Description Of Items");
            rowHead.createCell(3).setCellValue("Unit");
            rowHead.createCell(4).setCellValue("Rate");
            rowHead.createCell(5).setCellValue("Quantity");
            rowHead.createCell(6).setCellValue("Amount");

            sheet.setColumnWidth(1, 15 * 256);
            sheet.setColumnWidth(2, 60 * 256);
            sheet.setColumnWidth(4, 15 * 256);
            sheet.setColumnWidth(5, 15 * 256);
            sheet.setColumnWidth(6, 15 * 256);

            List<SubModel> models = getSubFile();

            SubModel subModel = models.get(1);
            HashMap<String, Double> hash = subModel.getConstituentSheets();

            int x = 7;

            for (Map.Entry mapElement : hash.entrySet()) {

                String key = (String) mapElement.getKey();
                rowHead.createCell(x).setCellValue("Quantity " + "[ " + key + " ]");
                sheet.setColumnWidth(x, 15 * 256);
                x++;
            }

            int i = 2;
            for (SubModel model : models) {
                XSSFRow row = sheet.createRow((short) i);

                row.createCell(1).setCellValue(model.getSOR());
                row.createCell(5).setCellValue(model.getQuantity());
                row.createCell(6).setCellValue(model.getAmount());

                //    System.out.println(model.getStatus());

                if (!model.getStatus())  //item not present in main db
                {
                    row.createCell(2).setCellValue(model.getDESCRIPTION_OF_ITEMS());

                    CellStyle style = workbook.createCellStyle();
                    style.setFillBackgroundColor(IndexedColors.SKY_BLUE.getIndex());
                    style.setFillPattern(FillPatternType.DIAMONDS);
                    row.setRowStyle(style);
                }


                if (searchRate(model.getSOR()).isPresent()) {  //item present in main db

                    Model model1 = searchRate(model.getSOR()).get();
                    row.createCell(2).setCellValue(model1.getDESCRIPTION_OF_ITEMS());
                    row.createCell(3).setCellValue(model1.getUnit());
                    row.createCell(4).setCellValue(model1.getRate());

                }

                HashMap<String, Double> hm = model.getConstituentSheets();

                x = 7;

                for (Map.Entry mapElement : hm.entrySet()) {

                    double value = ((double) mapElement.getValue());

                    row.createCell(x).setCellValue(value);
                    x++;
                }

                i++;
            }

            HttpHeaders header = new HttpHeaders();
            header.setContentType(new MediaType("application", "force-download"));
            header.set(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=FinalResult.xlsx");
            workbook.write(stream);
            workbook.close();

            return new ResponseEntity<>(new ByteArrayResource(stream.toByteArray()),
                    header, HttpStatus.CREATED);

        } catch (Exception e) {
            e.printStackTrace();
            return new ResponseEntity<>(HttpStatus.INTERNAL_SERVER_ERROR);
        }

    }
}