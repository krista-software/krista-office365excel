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

import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.Drive;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.ExcelFile;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProviderFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.KristaMediaClient;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.DataModification;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.DriveUtil;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.Result;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.StringUtil;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import com.microsoft.graph.http.GraphServiceException;
import com.microsoft.graph.models.DriveItem;
import com.microsoft.graph.models.RemoteItem;
import com.microsoft.graph.options.QueryOption;
import com.microsoft.graph.requests.*;
import org.jetbrains.annotations.Nullable;
import org.jvnet.hk2.annotations.Service;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

@Service
public class DriveService {

    private final GraphServiceClientProvider clientProvider;
    private static final Logger LOGGER = LoggerFactory.getLogger(DriveService.class);
    private final KristaMediaClient kristaMediaClient;


    @Inject
    public DriveService(GraphServiceClientProviderFactory providerFactory, KristaMediaClient kristaMediaClient) {
        this.clientProvider = providerFactory.create();
        this.kristaMediaClient = kristaMediaClient;
    }

    public List<Drive> getAllSharedDrivesInformation() {
        List<DriveItem> currentDriveItems = getDriveItemsSharedWithMe();
        if (currentDriveItems == null) {
            return Collections.emptyList();
        }
        return Result.getDrives(currentDriveItems);
    }

    public List<ExcelFile> getSharedExcelFilesInformation(String fileName) {
        List<DriveItem> currentDriveItems = getDriveItemsSharedWithMe();
        return currentDriveItems == null ? null : ExcelService.getExcelFiles(fileName, currentDriveItems);
    }

    private List<DriveItem> getDriveItemsSharedWithMe() {
        try {
            DriveSharedWithMeCollectionPage sharedWithMeCollectionPage = clientProvider.getDriveRequestBuilder(null).sharedWithMe()
                    .buildRequest()
                    .top(2000)
                    .get();

            List<DriveItem> currentDriveItems = new ArrayList<>();
            while (sharedWithMeCollectionPage != null) {
                currentDriveItems.addAll(sharedWithMeCollectionPage.getCurrentPage());
                DriveSharedWithMeCollectionRequestBuilder nextPage = sharedWithMeCollectionPage.getNextPage();
                if (nextPage != null) {
                    sharedWithMeCollectionPage = nextPage.buildRequest().get();
                } else {
                    break;
                }
            }
            return currentDriveItems;
        } catch (GraphServiceException cause) {
            LOGGER.error("Failed to get drive item: {}", cause.getMessage(), cause);
            return null;
        }
    }

    public List<ExcelFile> getExcelFileFromSharedDrive(String sharedFolderName, String sharedFileName) {

        if (sharedFolderName == null || sharedFolderName.isBlank()) {
            List<ExcelFile> sharedExcelFilesInformation = getSharedExcelFilesInformation(sharedFileName);
            if (sharedExcelFilesInformation != null && sharedExcelFilesInformation.isEmpty()) {
                LOGGER.info("Shared file {} is not accessible.", sharedFileName);
            }
            return sharedExcelFilesInformation;
        }

        String[] folderPath = sharedFolderName.split(Constants.BACKWARD_SLASH);
        List<DriveItem> driveItemsSharedWithMe = getDriveItemsSharedWithMe();
        if (driveItemsSharedWithMe == null) {
            return Collections.emptyList();
        }

        DriveItem firstDriveItemFolder = null;
        // Check folder is available at first level
        for (DriveItem driveItem : driveItemsSharedWithMe) {
            if ((driveItem.folder != null && driveItem.name != null && driveItem.name.equalsIgnoreCase(folderPath[0]))) {
                // As we get folder name, no need to check other DriveItem information
                // TODO: Need to cover, multiple folder "Demo" shared by multiple people

                // Folder "Demo" shared by single person
                firstDriveItemFolder = driveItem;
                break;
            }
        }

        if (firstDriveItemFolder == null) {
            LOGGER.info("Shared folder {} is not accessible", folderPath[0]);
            return Collections.emptyList();
        }

        if (folderPath.length == 1) {
            return getFileFromFirstLevelSharedFolder(sharedFileName, firstDriveItemFolder);
        }
        return getFileFromNestedSharedFolder(sharedFileName, folderPath, firstDriveItemFolder);

    }

    private List<ExcelFile> getFileFromNestedSharedFolder(String sharedFileName, String[] folderPath, DriveItem firstDriveItemFolder) {
        if (firstDriveItemFolder == null || firstDriveItemFolder.remoteItem == null || firstDriveItemFolder.remoteItem.parentReference == null) {
            return null;
        }

        List<ExcelFile> excelFiles = new ArrayList<>();
        boolean folderAvailable = false;
        DriveItem currentFolder = firstDriveItemFolder;
        String driveId = firstDriveItemFolder.remoteItem.parentReference.driveId;
        List<DriveItem> folderItems = null;

        for (int i = 1; i < folderPath.length + 1; i++) {
            DriveRequestBuilder driveRequestBuilder1 = clientProvider.getDriveRequestBuilder(driveId);
            folderItems = getFolderItems(driveRequestBuilder1, currentFolder.id);
            if (folderItems != null) {
                for (DriveItem driveItem : folderItems) {
                    if ((driveItem.folder != null && driveItem.name != null && driveItem.name.equalsIgnoreCase(folderPath[i]))) {
                        folderAvailable = true;
                        currentFolder = driveItem;
                        break;
                    }
                }
            }
            if (!folderAvailable) {
                LOGGER.info("Shared folder {} is not accessible.", folderPath[i]);
                break;
            }
        }

        if (folderAvailable) {
            excelFiles = getExcelFileFromSharedFolder(sharedFileName, folderItems);
            if (excelFiles == null) {
                throw new IllegalArgumentException("Shared file is not accessible");
            }
            if (excelFiles.isEmpty()) {
                LOGGER.info("Shared file {} is not accessible.", sharedFileName);
            }
        }
        return excelFiles;
    }

    @Nullable
    private List<ExcelFile> getFileFromFirstLevelSharedFolder(String sharedFileName, DriveItem firstDriveItemFolder) {
        List<ExcelFile> excelFiles = new ArrayList<>();
        // If only one folder name is given and that is available at first level
        if (firstDriveItemFolder != null) {
            RemoteItem remoteItem = firstDriveItemFolder.remoteItem;
            if (remoteItem != null && remoteItem.parentReference != null) {
                DriveRequestBuilder driveRequestBuilder1 = clientProvider.getDriveRequestBuilder(remoteItem.parentReference.driveId);
                List<DriveItem> folderItems = getFolderItems(driveRequestBuilder1, firstDriveItemFolder.id);
                excelFiles = getExcelFileFromSharedFolder(sharedFileName, folderItems);
            }
        } else {
            // If any folder name is not available, then path is invalid
            return null;
        }
        if (excelFiles == null) {
            throw new IllegalArgumentException("Shared File is not accessible");
        }
        if (excelFiles.isEmpty()) {
            LOGGER.info("Shared file {} is not accessible.", sharedFileName);
        }
        return excelFiles;
    }

    private List<DriveItem> getFolderItems(DriveRequestBuilder driveRequestBuilder, String folderId) {
        try {
            DriveItemCollectionPage children = driveRequestBuilder.items(folderId).children().buildRequest().get();

            List<DriveItem> currentDriveItems = new ArrayList<>();
            while (children != null) {
                currentDriveItems.addAll(children.getCurrentPage());
                DriveItemCollectionRequestBuilder nextPage = children.getNextPage();
                if (nextPage != null) {
                    children = nextPage.buildRequest().get();
                } else {
                    break;
                }
            }

            return currentDriveItems;
        } catch (GraphServiceException cause) {
            LOGGER.error("Failed to get folder items: {}", cause.getMessage(), cause);
            return null;
        }
    }

    private List<ExcelFile> getExcelFileFromSharedFolder(String fileName, List<DriveItem> currentDriveItems) {
        if (StringUtil.isNullOrEmpty(fileName) && currentDriveItems == null) {
            return null;
        }

        List<ExcelFile> excelFiles = new ArrayList<>();

        for (DriveItem driveItem : currentDriveItems) {
            if (driveItem.name != null && driveItem.name.equalsIgnoreCase(fileName)) {
                String fileDriveId = driveItem.parentReference == null ? null :
                        driveItem.parentReference.driveId == null ? null : driveItem.parentReference.driveId;
                excelFiles.add(new ExcelFile(driveItem.id, fileDriveId, driveItem.name, driveItem.webUrl));
            }
        }
        return excelFiles;
    }

    public app.krista.model.base.File getPdfFile(String fileId, String driveId) throws IOException {
        if (fileId == null) {
            LOGGER.info("Failed to get Pdf file. Given Id is invalid.");
            return null;
        }

        try {

            InputStream inputStream = getInputStreamForExcelFile(fileId, driveId);

            String fileName = "doc_pdf_" + UUID.randomUUID() + ".pdf";
            java.io.File pdfFile = DriveUtil.copyInputStreamToFile(inputStream, fileName);
            if (pdfFile == null) {
                throw new IllegalArgumentException("Failed to create PDF file.");
            }
            return kristaMediaClient.toKristaFile(pdfFile);
        } catch (IOException | GraphServiceException cause) {
            LOGGER.info("Failed to get PDF file {}", cause.getMessage(), cause);
            return null;
        }

    }

    @Nullable
    private InputStream getInputStreamForExcelFile(String fileId, String driveId) {
        LinkedList<QueryOption> requestOptions = new LinkedList<>();
        requestOptions.add(new QueryOption("format", "pdf"));

        return clientProvider.getDriveRequestBuilder(driveId).items(fileId).content().buildRequest(requestOptions).get();
    }

    public Map<String, Object> getPdfFileForGivenExcel(app.krista.model.base.File excelFile, String path) {


        if (excelFile == null || excelFile.getExtension() == null || !excelFile.getExtension().equalsIgnoreCase("xlsx")) {
            LOGGER.info("Failed to get PDF File. Upload Excel File.");
            return Collections.emptyMap();
        }

        try {
            String pdfFileId = uploadFileToDrive(kristaMediaClient, excelFile, path);
            app.krista.model.base.File pdfFile = getPdfFile(pdfFileId, null);

            return DataModification.getResultMap(pdfFileId, pdfFile);
        } catch (IOException | GraphServiceException cause) {
            LOGGER.info("Failed to get PDF file {}", cause.getMessage(), cause);
            return Map.of();
        }
    }

    private String uploadFileToDrive(KristaMediaClient kristaMediaClient, app.krista.model.base.File file, String path) throws IOException {
        java.io.File parsedFile = kristaMediaClient.toJavaFile(file);
        String finalPath = path == null ? parsedFile.getName() : path + "/" + parsedFile.getName();
        final DriveItem driveItem = clientProvider.getDriveRequestBuilder(null).root().itemWithPath(finalPath).content()
                .buildRequest()
                .put(DriveUtil.getByteArray(parsedFile));
        return driveItem != null ? driveItem.id : null;
    }

}
