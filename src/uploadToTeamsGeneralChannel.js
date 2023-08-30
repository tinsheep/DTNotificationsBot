
const { basename } = require('path');
const { readFile } = require('fs').promises;
const { FileUpload, OneDriveLargeFileUploadTask } = require('@microsoft/microsoft-graph-client');


async function uploadToTeamsGeneralChannel(driveItemId, filePath, graphClient) {
    try {
        const file = await readFile(filePath);
        const fileName = basename(filePath);

        const options = {
            path: 'General',
            fileName: fileName,
            rangeSize: 1024 * 1024,
            uploadSessionURL: `/drives/${driveItemId}/root:/General/${fileName}:/createUploadSession`,            
            uploadEventHandlers: {
                progress: (range, _) => {
                    console.log(`Uploaded bytes ${range?.minValue} to ${range?.maxValue}`);
                },
            }
        };

        const fileUpload = new FileUpload(file, fileName, file.byteLength);
        const uploadTask = await OneDriveLargeFileUploadTask.createTaskWithFileObject(
            graphClient,
            fileUpload,
            options
        );

        const uploadResult = await uploadTask.upload();
        const driveItem = uploadResult.responseBody;

        if (driveItem && driveItem.hasOwnProperty('id')) {
            console.log(`Uploaded file with ID: ${driveItem.id}`);
            const itemUrl = `https://teams.microsoft.com/_#/files/tab/${driveItem.parentReference.driveId}/${driveItem.parentReference.path}/children/${driveItem.name}/view`;
            return itemUrl;
        }

        console.log(`Uploaded file with ID: ${driveItem?.id}`);
    } catch (error) {
        console.error(`Error uploading file: ${error.message}`);
    }
}

module.exports = {
    uploadToTeamsGeneralChannel
};
