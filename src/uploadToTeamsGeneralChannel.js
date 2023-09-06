
const { basename } = require('path');
const { readFile } = require('fs').promises;
const { FileUpload, OneDriveLargeFileUploadTask } = require('@microsoft/microsoft-graph-client');


async function uploadToTeamsGeneralChannel(driveItemId, filePath, graphClient, teamId, channelId) {
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
            // call a function to create a deep link to the file
            const itemUrl = await createDeepLink(driveItem, teamId, channelId);
            return itemUrl;
        }

        console.log(`Uploaded file with ID: ${driveItem?.id}`);
    } catch (error) {
        console.error(`Error uploading file: ${error.message}`);
    }
}

async function createDeepLink(driveItem, teamId, channelId) {
    try {
        // Extract the GUID from the eTag property of the driveItem object
        const fileId = driveItem.eTag.match(/"({?[a-z0-9\-]+}?),\d+"/i)[1];
        const objectUrl = driveItem.webUrl;

        // Get the fileType from the objectUrl
        const fileType = objectUrl.match(/\.([^.]+)$/)[1];

        // Construct the baseUrl from the objectUrl
        const baseUrl = objectUrl.replace(/\/Shared%20Documents.*/, '');
        const itemUrl = 'https://teams.microsoft.com/l/file/' + fileId + '?tenantId=' + process.env.TEAMS_APP_TENANT_ID + '&fileType=' + fileType + '&objectUrl=' + encodeURIComponent(objectUrl) + '&baseUrl=' + encodeURIComponent(baseUrl) + '&serviceName=teams&threadId=' + channelId + '&groupId=' + teamId;

        // Return the itemUrl
        return itemUrl;
    } catch (error) {
        console.error(`Error creating deep link: ${error.message}`);
    }
}

module.exports = {
    uploadToTeamsGeneralChannel
};
