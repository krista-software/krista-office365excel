/*
 * Office365excel Extension for Krista
 * Copyright (C) 2025 Krista Software
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 */

package app.krista.extensions.essentials.spreadsheets.office365excel.service;

import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.ExcelSheet;
import app.krista.extensions.essentials.spreadsheets.office365excel.graph.v2.WorkbookRangeClearRequestBuilderV2;
import app.krista.extensions.essentials.spreadsheets.office365excel.graph.v2.WorkbookWorksheetRangeRequestBuilderV2;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProviderFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.Result;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.StringUtil;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.Worksheets;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.parser.DataParser;
import app.krista.model.base.File;
import app.krista.model.base.FreeForm;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.microsoft.graph.http.GraphServiceException;
import com.microsoft.graph.models.*;
import com.microsoft.graph.requests.*;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;
import org.jvnet.hk2.annotations.Service;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.io.IOException;
import java.util.List;
import java.util.*;

@Service
public class WorksheetService {

    public static final String TEXT = "Text";
    private final GraphServiceClientProvider clientProvider;
    private final WorkbookOperation workbookOperation;
    private static final Logger LOGGER = LoggerFactory.getLogger(WorksheetService.class);
    private final StringUtil stringUtil;


    @Inject
    public WorksheetService(WorkbookOperation workbookOperation, GraphServiceClientProviderFactory providerFactory, StringUtil stringUtil) {
        this.workbookOperation = workbookOperation;
        clientProvider = providerFactory.create();
        this.stringUtil = stringUtil;
    }

    private WorkbookWorksheetRequestBuilder worksheetRequest(String driveId, String fileId, String sheetName) {
        return clientProvider.getDriveRequestBuilder(driveId).items(fileId).workbook().worksheets(sheetName);
    }

    /**
     * Retrieves all worksheets from an Excel workbook with pagination support.
     * Iterates through all pages of results to return a complete list of worksheets.
     *
     * @param workbookId the ID of the Excel workbook
     * @param driveId the ID of the drive containing the workbook
     * @param accountId the account ID for authentication (null to use default)
     * @param invokeAsUser true to invoke as user context, false for service account
     * @return list of ExcelSheet objects containing worksheet information, empty list if none found or on error
     */
    public List<ExcelSheet> getWorksheets(String workbookId, String driveId, String accountId, Boolean invokeAsUser) {
        List<WorkbookWorksheet> worksheets = new ArrayList<>();
        WorkbookWorksheetCollectionPage worksheetCollectionPage = listWorksheets(workbookId, driveId, accountId, invokeAsUser);
        if (worksheetCollectionPage == null) {
            return List.of();
        }

        while (worksheetCollectionPage != null) {
            worksheets.addAll(worksheetCollectionPage.getCurrentPage());
            WorkbookWorksheetCollectionRequestBuilder nextPage = worksheetCollectionPage.getNextPage();
            if (nextPage != null) {
                worksheetCollectionPage = nextPage.buildRequest().get();
            } else {
                break;
            }
        }
        return Worksheets.getWorksheets(worksheets);
    }

    /**
     * Lists worksheets from a workbook using the Microsoft Graph API.
     *
     * @param workbookId the ID of the Excel workbook
     * @param driveId the ID of the drive containing the workbook
     * @param accountId the account ID for authentication
     * @param invokeAsUser true to invoke as user context, false for service account
     * @return WorkbookWorksheetCollectionPage containing the first page of worksheets, null on error
     */
    private WorkbookWorksheetCollectionPage listWorksheets(String workbookId, String driveId, String accountId, Boolean invokeAsUser) {
        try {
            return clientProvider.getDriveRequestBuilder(driveId, accountId, invokeAsUser)
                    .items(workbookId)
                    .workbook()
                    .worksheets()
                    .buildRequest()
                    .get();

        } catch (GraphServiceException | IOException cause) {
            LOGGER.error("Failed to list worksheets for workbook {}: {}", workbookId, cause.getMessage(), cause);
            return null;
        }
    }

    /**
     * Deletes a worksheet from an Excel workbook.
     *
     * @param workbookID the ID of the Excel workbook
     * @param worksheetName the name of the worksheet to delete
     * @return success message indicating the worksheet was deleted or already deleted
     * @throws IllegalStateException if deletion fails due to permissions, protection, or other errors
     */
    public String deleteWorksheets(String workbookID, String worksheetName) {
        try {
            worksheetRequest(null, workbookID, worksheetName).buildRequest().delete();
            return worksheetName + " Deleted Successfully";
        } catch (GraphServiceException cause) {
            if (cause.getResponseCode() == 404) {
                return worksheetName + " already deleted";
            }
            String errorMessage = cause.getServiceError() != null ? cause.getServiceError().message : cause.getMessage();
            LOGGER.error("Failed to delete worksheet '{}': {}", worksheetName, errorMessage, cause);
            throw new IllegalStateException("Unable to delete worksheet '" + worksheetName + "': " + errorMessage + ". " +
                    "Please check that you have permission to modify this file and the worksheet is not protected or in use.", cause);
        }
    }

    public Boolean deleteTableColumn(String workbookID, String tableName, String columnName) {
        try {
            getWorkbookTableColumn(workbookID, tableName, columnName, null);
            return true;
        } catch (GraphServiceException cause) {
            String errorMessage = cause.getServiceError() != null ? cause.getServiceError().message : cause.getMessage();
            LOGGER.error("Failed to delete column '{}' from table '{}': {}", columnName, tableName, errorMessage, cause);
            return false;
        }
    }

    /*
       Adding this method to have the backWard compatibility for "deleteTableColumn" method used in table Area
     */
    public Boolean deleteTableColumnInSharedFile(String workbookID, String tableName, String columnName, String driveId) {
        try {
            getWorkbookTableColumn(workbookID, tableName, columnName, driveId);
            return true;
        } catch (GraphServiceException cause) {
            String errorMessage = cause.getServiceError() != null ? cause.getServiceError().message : cause.getMessage();
            LOGGER.error("Failed to delete column '{}' from table '{}' in shared file: {}", columnName, tableName, errorMessage, cause);
            return false;
        }
    }

    @Nullable
    private WorkbookTableColumn getWorkbookTableColumn(String workbookID, String tableName, String columnName, String driveId) {
        return this.clientProvider.getDriveRequestBuilder(driveId).items(workbookID).workbook().tables(tableName).columns(columnName)
                .buildRequest()
                .delete();
    }

    public List<FreeForm> getRows(String driveId, String workbookID,
                                  String worksheetName, String range) {
        try {
            return getDataForRange(driveId, workbookID, worksheetName, range);
        } catch (GraphServiceException cause) {
            String errorMessage = cause.getServiceError() != null ? cause.getServiceError().message : cause.getMessage();
            LOGGER.error("Failed to get rows from worksheet '{}' with range '{}': {}", worksheetName, range, errorMessage, cause);
            return Collections.emptyList();
        }
    }

    public Map<String, Object> findMinMaxValueFromRowDataByColumnRange(String driveId, String workbookId, String workSheetName, String keyColumnName, String columnRange) {

        if (workbookId == null || workbookId.isBlank()) {
            LOGGER.warn("Cannot find min/max values: Workbook ID is missing or empty");
            return Collections.emptyMap();
        } else if (workSheetName == null || workSheetName.isBlank()) {
            LOGGER.warn("Cannot find min/max values: Worksheet name is missing or empty");
            return Collections.emptyMap();
        } else if (keyColumnName == null || keyColumnName.isBlank()) {
            LOGGER.warn("Cannot find min/max values: Key column name is missing or empty");
            return Collections.emptyMap();
        } else if (columnRange == null || columnRange.isBlank()) {
            LOGGER.warn("Cannot find min/max values: Column range is missing or empty");
            return Collections.emptyMap();
        }

        try {
            WorkbookRange workbookRange = getWorkbookRange(driveId, workbookId, workSheetName);
            if (workbookRange == null || workbookRange.columnCount == null) {
                LOGGER.error("Workbook range is null");
                return Collections.emptyMap();
            }
            String[] sequencedAlphabetsArr = Worksheets.getSequencedAlphabetsForColumns(workbookRange.columnCount);
            String[] columnRanges = columnRange.trim().split(Constants.COLON);

            int keyColumnIndex = Arrays.asList(sequencedAlphabetsArr).indexOf(keyColumnName);
            int startRangeColumnIndex = Arrays.asList(sequencedAlphabetsArr).indexOf(columnRanges[0]);
            int endRangeColumnIndex = Arrays.asList(sequencedAlphabetsArr).indexOf(columnRanges[1]);
            if (parseColumnRange(keyColumnName, columnRange, keyColumnIndex, startRangeColumnIndex, endRangeColumnIndex)) {
                return null;
            }
            JsonArray worksheetData = (JsonArray) workbookRange.text;
            if (worksheetData == null) {
                LOGGER.error("Worksheet data is null");
                return Collections.emptyMap();
            }
            JsonArray valueTypes = workbookRange.valueTypes.getAsJsonArray();
            JsonArray dataFormats = workbookRange.numberFormat.getAsJsonArray();
            List<Object> keyColumnData = Worksheets.createKeyColumn(worksheetData, valueTypes, dataFormats, keyColumnIndex);
            List<JsonArray> rangeWorksheetData = Worksheets.getWorksheetDataFromRange(worksheetData, startRangeColumnIndex, endRangeColumnIndex);
            Map<String, Double> minMaxValueInformation = Worksheets.findMinMaxValueInWorksheetRange(rangeWorksheetData);


            return Result.getResultFromMinMaxValueInformation(keyColumnData, minMaxValueInformation);
        } catch (GraphServiceException cause) {
            LOGGER.error("Failed to find min and max values: {}", cause.getMessage(), cause);
            return Collections.emptyMap();
        }
    }

    @Nullable
    public WorkbookRange getWorkbookRange(String driveId, String workbookId, String workSheetName) {

        return clientProvider.getDriveRequestBuilder(driveId).items(workbookId).workbook().worksheets(workSheetName).usedRange()
                .buildRequest()
                .get();
    }

    private static boolean parseColumnRange(String keyColumnName, String columnRange, int keyColumnIndex, int startRangeColumnIndex, int endRangeColumnIndex) {
        if (startRangeColumnIndex > endRangeColumnIndex || endRangeColumnIndex == -1 || startRangeColumnIndex == -1) {
            LOGGER.warn("Invalid column range '{}': Start column must come before end column and both must exist in the worksheet", columnRange);
            return true;
        }
        if (keyColumnIndex == -1) {
            LOGGER.warn("Column '{}' not found in worksheet. Column names are case-sensitive.", keyColumnName);
            return true;
        }
        return false;
    }

    /**
     * range can be sheetName!cellRange, example1-> Sheet1!A1:E7 or Sheet1!TableName
     *
     */
    private List<FreeForm> getDataForRange(String driveId, String workbookID,
                                           String worksheetName, String range) {
        FreeForm errorResult = new FreeForm();
        try {
            WorkbookWorksheetRangeParameterSet parameterSet = getRangeParameterSet(range);

            WorkbookWorksheetRangeRequestBuilder baseBuilder = worksheetRequest(driveId, workbookID, worksheetName).range(parameterSet);
            if (baseBuilder.getRequestUrl() == null || baseBuilder.getClient() == null) {
                LOGGER.error("Base builder is null");
                return null;
            }
            WorkbookWorksheetRangeRequestBuilderV2 builderV2 = new WorkbookWorksheetRangeRequestBuilderV2(
                    baseBuilder.getRequestUrl(), baseBuilder.getClient(), baseBuilder.getOptions(), parameterSet);

            WorkbookRange workbookRange = builderV2.buildRequest().get();
            if (workbookRange == null || workbookRange.text == null) {
                LOGGER.error("Workbook range is null");
                return null;
            }
            JsonElement jsonElement = workbookRange.text;
            return DataParser.parse(jsonElement);
        } catch (GraphServiceException graphServiceException) {
            errorResult.put("Error Code", "Text", graphServiceException.getError() != null &&
                    graphServiceException.getError().error != null ?
                    graphServiceException.getError().error.code : "NOT FOUND");
            errorResult.put("Error Message", "Text",
                    "The worksheet '" + worksheetName + "' was not found in this workbook. " +
                            "Please check the sheet name for typos and ensure there are no extra spaces. " +
                            "You can use 'Get Sheets Information' to see all available worksheets.");
            return List.of(errorResult);
        }
    }

    @NotNull
    public static WorkbookWorksheetRangeParameterSet getRangeParameterSet(String range) {
        return WorkbookWorksheetRangeParameterSet.newBuilder().withAddress(range).build();
    }

    /**
     * Retrieves data from a single cell in an Excel worksheet.
     *
     * @param driveId the ID of the drive containing the workbook
     * @param workbookID the ID of the Excel workbook
     * @param worksheetName the name of the worksheet
     * @param cellID the cell address (e.g., "A1", "B5")
     * @return FreeForm object containing the cell data, empty FreeForm on error
     */
    public FreeForm getCellData(String driveId, String workbookID,
                                String worksheetName, String cellID) {
        try {
            List<FreeForm> results = getDataForRange(driveId, workbookID, worksheetName, cellID + ":" + cellID);
            return results == null || results.isEmpty() ? new FreeForm() : results.getFirst();
        } catch (GraphServiceException cause) {
            LOGGER.error("Failed to get cell data: {}", cause.getMessage(), cause);
            return new FreeForm();
        }
    }

    /**
     * Creates a new worksheet in an Excel workbook.
     *
     * @param workbookId the ID of the Excel workbook
     * @param newWorksheetName the name for the new worksheet
     * @return ExcelSheet object containing the new worksheet information, empty ExcelSheet on error
     */
    public ExcelSheet createWorksheet(String workbookId, String newWorksheetName) {
        try {
            WorkbookWorksheetAddParameterSet worksheetAddParameterSet = WorkbookWorksheetAddParameterSet.newBuilder().withName(newWorksheetName).build();

            WorkbookWorksheet workbookWorksheet = clientProvider.getDriveRequestBuilder(null).items(workbookId).workbook()
                    .worksheets()
                    .add(worksheetAddParameterSet)
                    .buildRequest().post();
            if (workbookWorksheet != null) {
                return Worksheets.getWorksheet(workbookWorksheet);
            }
            return new ExcelSheet("", "");
        } catch (GraphServiceException cause) {
            LOGGER.error("Failed to create worksheet: {}", cause.getMessage(), cause);
            return new ExcelSheet("", "");
        }
    }

    /**
     * Clears all content from a specified range of cells in a worksheet.
     * Clears values, formatting, and formulas from the specified range.
     *
     * @param workbookId the ID of the Excel workbook
     * @param worksheetName the name of the worksheet
     * @param cellRange the range of cells to clear (e.g., "A1:B10")
     * @return true if cells were cleared successfully, false on error
     */
    public Boolean clearWorksheetCells(String workbookId, String worksheetName, String cellRange) {
        try {
            WorkbookWorksheetRangeParameterSet worksheetRangeParameterSet = getRangeParameterSet(cellRange);

            WorkbookWorksheetRangeRequestBuilder worksheetRangeRequestBuilder = worksheetRequest(null, workbookId, worksheetName)
                    .range(worksheetRangeParameterSet);
            if (worksheetRangeRequestBuilder.getRequestUrl() == null || worksheetRangeRequestBuilder.getClient() == null) {
                LOGGER.error("Worksheet range request builder is null");
                return false;
            }
            WorkbookWorksheetRangeRequestBuilderV2 rangeRequestBuilderV2 = new WorkbookWorksheetRangeRequestBuilderV2(
                    worksheetRangeRequestBuilder.getRequestUrl(), worksheetRangeRequestBuilder.getClient(),
                    worksheetRangeRequestBuilder.getOptions(), worksheetRangeParameterSet);

            WorkbookRangeClearParameterSet parameterSet = WorkbookRangeClearParameterSet
                    .newBuilder()
                    .withApplyTo(Constants.ApplyTo.All.name())
                    .build();

            if (rangeRequestBuilderV2.buildRequest().getRequestUrl() == null ||
                    rangeRequestBuilderV2.getClient() == null) {
                LOGGER.error("Range request builder V2 is null");
                return false;
            }
            WorkbookRangeClearRequest rangeClearRequest = WorkbookRangeClearRequestBuilderV2
                    .newBuilder(rangeRequestBuilderV2.buildRequest().getRequestUrl().toString(),
                            rangeRequestBuilderV2.getClient(), rangeRequestBuilderV2.getOptions(), parameterSet)
                    .buildRequest();
            rangeClearRequest.post();
            return true;
        } catch (GraphServiceException cause) {
            LOGGER.error("Failed to clear worksheet cells: {}", cause.getMessage(), cause);
            return false;
        }
    }

    public List<String> findValue(String driveId, String workbookId, String worksheetName, String lookupValue) {
        try {
            WorkbookRange wbRange = worksheetRequest(driveId, workbookId, worksheetName)
                    .usedRange()
                    .buildRequest()
                    .get();

            return Worksheets.search(wbRange, lookupValue);
        } catch (GraphServiceException cause) {
            LOGGER.error("Failed to find value: {}", cause.getMessage(), cause);
            return Collections.emptyList();
        }
    }

    public FreeForm addDataToCell(String driveId, String workbookID, String sheetName,
                                  String cellID, String dataToBeAdded) {
        try {
            updateCell(driveId, workbookID, sheetName, cellID, dataToBeAdded);
            FreeForm cellInformation = new FreeForm();
            cellInformation.put(Constants.CELL_ID, TEXT, cellID);
            cellInformation.put(Constants.VALUE, TEXT, dataToBeAdded);
            return cellInformation;
        } catch (GraphServiceException cause) {
            LOGGER.error("Failed to add data to cell: {}", cause.getMessage(), cause);
            return new FreeForm();
        }
    }

    public void updateCell(String driveId, String workbookID,
                           String sheetName, String cellID, String dataToBeAdded) {
        WorkbookWorksheetRangeParameterSet set = getRangeParameterSet(cellID);

        WorkbookWorksheetRangeRequestBuilder rangeRequestBuilder = worksheetRequest(
                driveId, workbookID, sheetName).range(set);
        if (rangeRequestBuilder.getRequestUrl() == null || rangeRequestBuilder.getClient() == null) {
            LOGGER.error("Range request is null");
            return;
        }
        WorkbookWorksheetRangeRequestBuilderV2 builderV2 = new WorkbookWorksheetRangeRequestBuilderV2(
                rangeRequestBuilder.getRequestUrl(), rangeRequestBuilder.getClient(), rangeRequestBuilder.getOptions(), set);

        WorkbookRange wbRange = new WorkbookRange();
        wbRange.address = sheetName + Constants.SHEET_RANGE_SEPARATOR + cellID;
        if (dataToBeAdded != null && dataToBeAdded.startsWith("=")) {
            wbRange.formulas = Constants.GSON.toJsonTree(dataToBeAdded);
        } else {
            wbRange.values = Constants.GSON.toJsonTree(dataToBeAdded);
        }
        builderV2.buildRequest().patch(wbRange);
    }

    public List<FreeForm> addRows(String driveId, String fileID, String sheetName,
                                  List<String> rows, List<File> rowsFromCsv, String address) throws IOException {
        List<String> rowsToAdd = stringUtil.getStringList(rows, rowsFromCsv);
        return workbookOperation.addRows(driveId, fileID, sheetName,
                StringUtil.appendCommas(rowsToAdd), address);
    }

    public List<FreeForm> addRows(String driveId, String fileID, String sheetName,
                                  List<String> rows, String address) throws IOException {
        return workbookOperation.addRowsForInsertAndUpdate(driveId, fileID, sheetName, rows, address);
    }

    public FreeForm getRow(String fileID, String SheetName, Double rowIndex, String driveId) {

        return workbookOperation.getRow(fileID, SheetName, rowIndex, driveId);

    }

    /**
     * Checks if a worksheet exists with the given file ID and worksheet name.
     *
     * @param fileId        The ID of the file containing the worksheet
     * @param worksheetName The name of the worksheet to check
     * @param driveId       The ID of the drive (null for personal drive)
     * @return true if the worksheet exists, false otherwise
     */
    public boolean worksheetExists(String fileId, String worksheetName, String driveId) {
        if (fileId == null || fileId.trim().isEmpty() || worksheetName == null || worksheetName.trim().isEmpty()) {
            return false;
        }

        try {
            LOGGER.info("Checking if worksheet '{}' exists in file '{}' (drive: '{}')", worksheetName, fileId, driveId);
            clientProvider.getDriveRequestBuilder(driveId).items(fileId).workbook().worksheets(worksheetName).buildRequest().get();
            LOGGER.info("Worksheet '{}' exists in file '{}' (drive: '{}')", worksheetName, fileId, driveId);
            return true;
        } catch (GraphServiceException cause) {
            if (cause.getResponseCode() == 404) {
                LOGGER.error("Worksheet '{}' does not exist in file '{}' (drive: '{}'): {}", worksheetName, fileId, driveId, cause.getMessage());
            } else {
                LOGGER.error("Error checking if worksheet '{}' exists in file '{}' (drive: '{}'): {}", worksheetName, fileId, driveId, cause.getMessage(), cause);
            }
            return false;
        } catch (Exception cause) {
            LOGGER.error("Unexpected error checking if worksheet '{}' exists in file '{}' (drive: '{}'): {}", worksheetName, fileId, driveId, cause.getMessage(), cause);
            return false;
        }
    }

    /**
     * Clears cells in a worksheet in a shared drive
     *
     * @param driveId       The ID of the shared drive
     * @param workbookId    The ID of the workbook
     * @param worksheetName The name of the worksheet
     * @param cellRange     The range of cells to clear
     * @return true if the operation was successful, false otherwise
     */
    public Boolean clearWorksheetCells(String driveId, String workbookId, String worksheetName, String cellRange) {
        try {
            WorkbookWorksheetRangeParameterSet worksheetRangeParameterSet = getRangeParameterSet(cellRange);

            WorkbookWorksheetRangeRequestBuilder worksheetRangeRequestBuilder = worksheetRequest(driveId, workbookId, worksheetName)
                    .range(worksheetRangeParameterSet);
            if (worksheetRangeRequestBuilder.getRequestUrl() == null || worksheetRangeRequestBuilder.getClient() == null) {
                LOGGER.error("Worksheet range request builder is null");
                return false;
            }
            WorkbookWorksheetRangeRequestBuilderV2 rangeRequestBuilderV2 = new WorkbookWorksheetRangeRequestBuilderV2(
                    worksheetRangeRequestBuilder.getRequestUrl(), worksheetRangeRequestBuilder.getClient(),
                    worksheetRangeRequestBuilder.getOptions(), worksheetRangeParameterSet);

            WorkbookRangeClearParameterSet parameterSet = WorkbookRangeClearParameterSet
                    .newBuilder()
                    .withApplyTo(Constants.ApplyTo.All.name())
                    .build();
            if (rangeRequestBuilderV2.buildRequest().getRequestUrl() == null ||
                    rangeRequestBuilderV2.getClient() == null) {
                LOGGER.error("Range request builder V2 is null");
                return false;
            }
            WorkbookRangeClearRequest rangeClearRequest = WorkbookRangeClearRequestBuilderV2
                    .newBuilder(rangeRequestBuilderV2.buildRequest().getRequestUrl().toString(),
                            rangeRequestBuilderV2.getClient(), rangeRequestBuilderV2.getOptions(), parameterSet)
                    .buildRequest();
            rangeClearRequest.post();
            return true;
        } catch (GraphServiceException cause) {
            LOGGER.error("Failed to clear worksheet cells in shared drive: {}", cause.getMessage(), cause);
            return false;
        }
    }

}
