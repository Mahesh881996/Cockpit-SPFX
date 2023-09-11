// import pnp and pnp logging system
import { spfi, SPFI, SPFx } from "@pnp/sp";
import { LogLevel, PnPLogging } from "@pnp/logging";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import "@pnp/sp/attachments";
import "@pnp/sp/lists/web";
import "@pnp/sp/site-users/web";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/files/folder";
import "@pnp/sp/security";
import "@pnp/sp/security/web";
import "@pnp/sp/site-users/web";
import "@pnp/sp/presets/all";
import { IAttachmentFileInfo } from "@pnp/sp/attachments";
import { WebPartContext } from "@microsoft/sp-webpart-base";

let _sp: SPFI = null;
export const getSP = (context?: WebPartContext): SPFI => {
    // if (_sp === null && context != null) {
        //You must add the @pnp/logging package to include the PnPLogging behavior it is no longer a peer dependency
        // The LogLevel set's at what level a message will be written to the console
        _sp = spfi().using(SPFx(context)).using(PnPLogging(LogLevel.Warning));
    // }
    return _sp;
};

export const addAttachments = async (files: IAttachmentFileInfo[], itemId: number, listName: string) => {
    let spcontext: SPFI = getSP();
    const item = await spcontext.web.lists.getByTitle(listName).items.getById(itemId);
    // const [batchedSP, execute] = spcontext.batched();
    files.map(async (file) => {
        await item.attachmentFiles.add(file.name, file.content);
    })
    return true;
    // await execute();
}

export const createListItem = async (listName: string, body: any, files?: IAttachmentFileInfo[]) => {
    let spcontext: SPFI = getSP();

    try {
        const item: any = await spcontext.web.lists.getByTitle(listName).items.add(body);
        if (files) {
            await addAttachments(files, item.Id, listName);
        }
        return item;
    }
    catch (err) {
        throw new Error(err);
    }
}
export const updateListItem = async (listName: string, body: any, id: number, files?: IAttachmentFileInfo[]) => {
    let spcontext: SPFI = getSP();

    try {
        let updateItem = await spcontext.web.lists
            .getByTitle(listName)
            .items
            .getById(id)
            .update(body).then(async (r: any) => {
                if (files) {
                    await addAttachments(files, id, listName);
                }
            });
        return updateItem;
    }
    catch (err) {
        throw new Error(err);
    }
}

export const ensureUser = async (loginName: string) => {
    let spcontext: SPFI = getSP();
    try {
        let userDetails = await spcontext.web.ensureUser(loginName);
        return userDetails;
    }
    catch (err) {
        throw new Error(err);
    }
}

export const getListItems = async (listName: string, select: string, lookup: string, filter: string, orderByField: string, context: any) => {
    let spcontext: SPFI = getSP(context);
    try {
        let allItems = orderByField === "" ? await spcontext.web.lists.getByTitle(listName).items.select(select).expand(lookup).filter(filter).top(5000)() : await spcontext.web.lists.getByTitle(listName).items.select(select).expand(lookup).filter(filter).orderBy(orderByField, true).top(5000)();
        return allItems;
    }
    catch (err) {
        throw new Error(err);
    }
}

export const getListItemById = async (listName: string, id: number, select: string, lookup: string) => {
    let spcontext: SPFI = getSP();
    try {
        let item = await spcontext.web.lists.getByTitle(listName).items.getById(id).select(select).expand(lookup)();
        return item;
    }
    catch (err) {
        throw new Error(err);
    }
}

export const getDocumentLibraryFiles = async (folderPath: string) => {
    let spcontext: SPFI = getSP();
    try {
        let allFiles = await spcontext.web.getFolderByServerRelativePath(folderPath).files();
        return allFiles;
    }
    catch (err) {
        throw new Error(err);
    }
}

export const getFileProperties = async (filePath: string, properties: any) => {
    let spcontext: SPFI = getSP();
    try {
        let fileProperties = await spcontext.web.getFileByServerRelativePath(filePath).getItem("Title", "Name", "ID", "Site", "Plant", "Area", "DocumentType", "DocumentNumber", "Discipline", "Created", "LinkFilename", "LiveOrProjectVersion");
        return fileProperties;
    }
    catch (err) {
        throw new Error(err);
    }
}

export const addFileToDocLib = async (path: string, fileName: string, fileContent: string | ArrayBuffer | Blob, properties: Record<string, any>, codes: Record<string, any>) => {
    let spcontext: SPFI = getSP();
    try {
        const file = await spcontext.web.getFolderByServerRelativePath(path).files.addUsingPath(fileName, fileContent, { Overwrite: true });
        const item: any = await file.file.getItem();
        switch (properties.Category) {
            case "LIVE":
                properties.DocumentNumber = "D-" + codes.siteCode + "-" + codes.plantCode + "-" + codes.typeCode + "-" + item.Id + "";
                properties["StatusCode"] = "1";
                break;
            case "PROJECT":
                properties.DocumentNumber = "PD-" + properties.ProjectNumber + "-" + codes.siteCode + "-" + codes.plantCode + "-" + codes.typeCode + "-" + item.Id + "";
                properties["StatusCode"] = "2";
                break;
        }
        properties["FileLeafRef"] = properties.DocumentNumber;
        await spcontext.web.lists.getByTitle("DocumentStatus").items.add({ Title: "Upload Document", DocId: item.Id, StatusCode: properties.StatusCode, AccessWFStatus: "In Progress", DeleteStatus: "NA", DocName: properties.DocumentNumber + '.' + fileName.split('.')[1], DocNumber: properties.DocumentNumber }).then(async (r: { data: { ID: number; }; }) => { console.log(r.data.ID) });
        delete properties.StatusCode;
        await item.update(properties);
        await spcontext.web.getFileByServerRelativePath(path + '/' + encodeURI(properties.DocumentNumber + '.' + file.data.Name.split('.')[1]) + '').checkin();
        return item;
    }
    catch (err) {
        throw new Error(err);
    }
}
export const updateDocProps = async (path: string, properties: Record<string, any>, codes: Record<string, any>, isCheckedOut: boolean, currentDocNo: string, newDocNo: string, documentLibPath: string) => {
    let spcontext: SPFI = getSP();
    try {
        if (!isCheckedOut) {
            await spcontext.web.getFileByServerRelativePath(path).checkout();
        }
        const item: any = await spcontext.web.getFolderByServerRelativePath(path).getItem();
        if (currentDocNo === newDocNo) {
            switch (properties.Category) {
                case "LIVE":
                    properties.DocumentNumber = "D-" + codes.siteCode + "-" + codes.plantCode + "-" + codes.typeCode + "-" + item.Id + "";
                    switch (properties.StatusCode) {
                        case "1":
                        case "2":
                            properties["StatusCode"] = "1";
                            break;
                    }
                    break;
                case "PROJECT":
                    properties.DocumentNumber = "PD-" + properties.ProjectNumber + "-" + codes.siteCode + "-" + codes.plantCode + "-" + codes.typeCode + "-" + item.Id + "";
                    switch (properties.StatusCode) {
                        case "1":
                        case "2":
                            properties["StatusCode"] = "2";
                            break;
                    }
                    break;
            }
        } else {
            properties.DocumentNumber = newDocNo;
            switch (properties.Category) {
                case "LIVE":
                    switch (properties.StatusCode) {
                        case "1":
                        case "2":
                            properties["StatusCode"] = "1";
                            break;
                    }
                    break;
                case "PROJECT":
                    switch (properties.StatusCode) {
                        case "1":
                        case "2":
                            properties["StatusCode"] = "2";
                            break;
                    }
                    break;
            }
        }
        properties["FileLeafRef"] = properties.DocumentNumber;
        const StatusRecord: any = await spcontext.web.lists.getByTitle("DocumentStatus").items.select("Title,Id,DocId").filter(`DocId eq ${item.Id}`).top(5000)();
        await spcontext.web.lists.getByTitle("DocumentStatus").items.getById(StatusRecord[0].Id).update({ Title: "Update Properties", StatusCode: properties.StatusCode, AccessWFStatus: properties.AccessWorkflowStatus, DocName: properties.DocumentNumber + '.' + path.split('.')[1], DocNumber: properties.DocumentNumber }).then(async (r: any) => { console.log(r) });
        delete properties.StatusCode;
        delete properties.AccessWorkflowStatus;
        await item.update(properties);
        await spcontext.web.getFileByServerRelativePath(documentLibPath + '/' + encodeURI(properties.DocumentNumber + '.' + path.split('.')[1]) + '').checkin();
        return item;
    }
    catch (err) {
        throw new Error(err);
    }
}
export const LogError = async (message: string, errorData: string, moduleName: string, methodName: string) => {
    let spcontext: SPFI = getSP();
    let user = await spcontext.web.currentUser()
    await spcontext.web.lists.getByTitle("ErrorLogs").items.add({
        MethodName: methodName,
        ModuleName: moduleName,
        ErrorMessage: message,
        ErrorDetails: JSON.stringify(errorData),
        LogLevel: 'Error',
        LoggedInUserName: user.Title,
        TimeStamp: new Date()
    });
}
