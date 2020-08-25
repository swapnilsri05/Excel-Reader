package in.hopscotch.ExcelReaderV1.controller;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import in.hopscotch.ExcelReaderV1.dto.ResponseModel;
import in.hopscotch.ExcelReaderV1.model.FileTemplate;
import in.hopscotch.ExcelReaderV1.utils.PoiUtils;

@RestController
public class ExcelReaderController {

    @PostMapping("/api/file/upload")
    public ResponseModel parseExcelAndReturnTemplate(@RequestParam("file") MultipartFile readExcelDataFile) throws IOException {
        ResponseModel response = new ResponseModel();
        List<FileTemplate> templateList = new ArrayList<>();
        XSSFWorkbook workbook = new XSSFWorkbook(readExcelDataFile.getInputStream());
        XSSFSheet worksheet = workbook.getSheetAt(1);
        for (int i = 1; i < worksheet.getPhysicalNumberOfRows(); i++) {
            FileTemplate template = new FileTemplate();
            XSSFRow row = worksheet.getRow(i);

            String value = PoiUtils.getStringCellValue(row, 0);
            if (!value.isEmpty()) {
                template.setStyleCode(value);
                template.setBrand(PoiUtils.getStringCellValue(row, 1));
                template.setVsku(PoiUtils.getStringCellValue(row, 2));
                template.setColor(PoiUtils.getStringCellValue(row, 3));
                template.setSize(PoiUtils.getStringCellValue(row, 4));
                template.setSizeInCms(PoiUtils.getStringCellValue(row, 5));
                template.setFromAge(PoiUtils.getStringCellValue(row, 6));
                template.setToAge(PoiUtils.getStringCellValue(row, 7));
                template.setGender(PoiUtils.getStringCellValue(row, 8));
                template.setReadyQuantity(PoiUtils.getStringCellValue(row, 9));
                template.setCost(PoiUtils.getStringCellValue(row, 10));
                template.setCustomCategory(PoiUtils.getStringCellValue(row, 11));
                template.setTopMaterial1(PoiUtils.getStringCellValue(row, 12));
                template.setTopMaterial1Composition(PoiUtils.getStringCellValue(row, 13));
                template.setTopMaterial2(PoiUtils.getStringCellValue(row, 14));
                template.setTopMaterial2Composition(PoiUtils.getStringCellValue(row, 15));
                template.setBottomMaterial1(PoiUtils.getStringCellValue(row, 16));
                template.setBottomMaterial1Composition(PoiUtils.getStringCellValue(row, 17));

                template.setBottomMaterial2(PoiUtils.getStringCellValue(row, 18));
                template.setBottomMaterial2Composition(PoiUtils.getStringCellValue(row, 19));

                template.setOuterMaterial1(PoiUtils.getStringCellValue(row, 20));
                template.setOuterMaterial1Composition(PoiUtils.getStringCellValue(row, 21));

                template.setOuterMaterial2(PoiUtils.getStringCellValue(row, 22));
                template.setOuterMaterial2Composition(PoiUtils.getStringCellValue(row, 23));

                template.setRoundWaist(PoiUtils.getStringCellValue(row, 24));
                template.setBottomLength(PoiUtils.getStringCellValue(row, 25));
                template.setRoundChest(PoiUtils.getStringCellValue(row, 26));
                template.setTopLength(PoiUtils.getStringCellValue(row, 27));
                template.setRoundChestOutwear(PoiUtils.getStringCellValue(row, 28));

                template.setTopLengthOutwear(PoiUtils.getStringCellValue(row, 29));

                template.setLength(PoiUtils.getStringCellValue(row, 30));


                template.setInsoleLength(PoiUtils.getStringCellValue(row, 31));
                

                template.setLegLength(PoiUtils.getStringCellValue(row, 32));
                
                template.setSellingCost(PoiUtils.getStringCellValue(row, 33));
                template.setDescriptionOrAccessoriesDescription(PoiUtils.getStringCellValue(row, 34));
                
                template.setProductName(PoiUtils.getStringCellValue(row, 35));
                template.setWhatIsIncluded(PoiUtils.getStringCellValue(row, 36));
                template.setPackingWeightInKgs(PoiUtils.getStringCellValue(row, 37));
                template.setPackingLength(PoiUtils.getStringCellValue(row, 38));
template.setPackingWidth(PoiUtils.getStringCellValue(row, 39));
template.setPackingHeight(PoiUtils.getStringCellValue(row, 40));
template.setPackingWeightInKgs(PoiUtils.getStringCellValue(row, 41));
                
                
                template.setRowCheck(PoiUtils.getStringCellValue(row, 42));
                template.setColumnToCheck(PoiUtils.getStringCellValue(row, 84));
                templateList.add(template);
            }

        }

        response.setData(templateList);
        response.setCode(0);
        return response;
    }

}
