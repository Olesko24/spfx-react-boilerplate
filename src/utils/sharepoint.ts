import { ISPFXContext, SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";
import "@pnp/sp/attachments";
import "@pnp/sp/profiles";
import "@pnp/sp/site-users";
import "@pnp/sp/fields";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import {
    IList,
    IListInfo,
} from "@pnp/sp/lists";
import { FieldTypes, IFieldInfo } from "@pnp/sp/fields";
import { IViewInfo } from "@pnp/sp/views/types";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/views";
import "@pnp/sp/items";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/items/get-all";
import "@pnp/sp/site-users";
import "@pnp/sp/regional-settings/web";

let sp: SPFI;

/**
 * Initializes the PnP SharePoint Object
 * @param context SP page context
 */
export async function initSP(context: ISPFXContext): Promise<void> {
    sp = spfi().using(SPFx(context));
    try {
        const siteAssetsList: IList = await sp.web.lists.ensureSiteAssetsLibrary();
        const title: IListInfo = await siteAssetsList.select("Title")();
        if (title) {
            console.log("[PnP SP] SharePoint Connection established!");
        }
    } catch (error) {
        throw new Error(`error while connecting to sharepoint: ${error}`);
    }
}
export async function getWebBaseUrl(): Promise<string> {
    try {
        const web = await sp.web.select("ServerRelativeUrl")();
        return web.ServerRelativeUrl;
    } catch (error) {
        throw new Error(`error while getting web base url: ${error}`);
    }
}

/**
 * Gets all List Items from a List.
 * @param listTitle List Name
 * @returns List Items
 */
export async function getListItems(
    listTitle: string,
    withAttachments?: boolean,
    queryFilter?: string,
    fields?: string[],
    lookup?: string[]
): Promise<any[]> {
    try {
        const query: any = sp.web.lists.getByTitle(listTitle).items;

        if (fields) {
            query.select(...fields);

            if (lookup) {
                query.expand(...lookup);
            }
        }

        if (queryFilter) {
            query.filter(queryFilter);
        }

        const items = await query.getAll();

        if (withAttachments) {
            for (const item of items) {
                item.Attachments = await sp.web.lists
                    .getByTitle(listTitle)
                    .items.getById(item.ID)
                    .attachmentFiles();
            }
        }

        return items;
    } catch (error) {
        throw new Error(
            `error while getting items from list ${listTitle}: ${error}`
        );
    }
}

/**
 * Creates a new Item in the specified List.
 * @param listTitle List Name
 * @param item Item Properties
 * @returns new created item
 */
export async function createListItem(listTitle: string, item: any) {
    try {
        const createdItem = await sp.web.lists
            .getByTitle(listTitle)
            .items.add({ ...item });
        return createdItem;
    } catch (error) {
        throw new Error(
            `error while creating new list item in list ${listTitle}: ${error}`
        );
    }
}

/**
 * Adds an Attachment to a List Item.
 * @param listTitle List Name
 * @param itemID Item ID
 * @param file File Content
 * @param fileName File Name
 */
export async function addAttachmentToListItem(listTitle: string, itemID: number, file: Blob | ArrayBuffer | string, fileName: string) {
    try {
        const attachment = await sp.web.lists
            .getByTitle(listTitle)
            .items.getById(itemID)
            .attachmentFiles.add(fileName, file);
        return attachment;
    } catch (error) {
        throw new Error(`error while adding attachment to list item: ${error}`);
    }
}

/**
 * Delete the attachment and send it to recycle bin
 * @param listTitle 
 * @param itemID 
 * @param fileName 
 */
export async function recycleAttachment(listTitle: string, itemID: number, fileName: string) {
    try {
        const attachment = await sp.web.lists
            .getByTitle(listTitle)
            .items.getById(itemID)
            .attachmentFiles.getByName(fileName)
            .recycle();
        return attachment;
    } catch (error) {
        throw new Error(`error while recycling attachment: ${error}`);
    }
}

/**
 * Returns the List Item by ID.
 * @param listTitle List Title
 * @param itemID Item ID
 * @returns list item
 */
export async function getListItemById(listTitle: string, itemID: number) {
    try {
        const item = await sp.web.lists
            .getByTitle(listTitle)
            .items.getById(itemID)();
        return item;
    } catch (error) {
        throw new Error(`error while getting list item by id ${itemID} : ${error}`);
    }
}

/**
 * Updates a specific Item by ID with a new Value.
 * @param listTitle List Title
 * @param itemID Item ID
 * @param updateValue new Item Properties
 * @returns updated item
 */
export async function updateListItemById(
    listTitle: string,
    itemID: number,
    updateValue: any
) {
    try {
        const updatedItem = await sp.web.lists
            .getByTitle(listTitle)
            .items.getById(itemID)
            .update(updateValue);
        return updatedItem;
    } catch (error) {
        throw new Error(
            `error while updating list item by id ${itemID} : ${error}`
        );
    }
}
export async function deleteListItemById(listTitle: string, itemID: number) {
    try {
        const deleteItem = await sp.web.lists
            .getByTitle(listTitle)
            .items.getById(itemID)
            .delete();
        return deleteItem;
    } catch (error) {
        throw new Error(
            `error while deleting list item by id ${itemID} : ${error}`
        );
    }
}

/**
 * Gets the current User properties.
 * @returns current user properties
 */
export async function getCurrentUserProfile() {
    try {
        //const profile = await sp.profiles.myProperties();
        const profile = await sp.web.currentUser();
        return profile;
    } catch (error) {
        throw new Error(`error while getting current user profile ${error}`);
    }
}

export async function getListFields(listTitle: string) {
    try {
        const fields = await sp.web.lists.getByTitle(listTitle).fields();
        return fields;
    } catch (err) {
        console.log(err);
    }
}
/**
 * Add a new List.
 * @param title List title
 * @param options options
 */
export async function addList(
    title: string,
    description?: string,
    template?: number
) {
    try {
        const list = await sp.web.lists.add(
            title,
            description,
            template
        );
        return list;
    } catch (error) {
        throw new Error(`error while creating the list ${error}`);
    }
}
/**
 * Ensure a list exists. If not it creates the list.
 * @param title List title
 * @param options options
 */
export async function ensureList(
    title: string,
    options?: any
) {
    try {
        const list = await sp.web.lists.ensure(title, ...options);
        return list;
    } catch (error) {
        throw new Error(`error while ensuring that list exists: ${error}`);
    }
}
/**
 * Add a new Field to a given List.
 * @param listname List title
 * @param title Field title
 * @param type Field type
 * @param options Field options
 */
export async function addFieldToListv2(
    listname: string,
    title: string,
    type: "text" | "number" | "user" | "boolean" | string,
    options?: any
) {
    try {
        let field: Partial<IFieldInfo>;
        if (!options) options = {};
        switch (type) {
            case "number":
                field = await sp.web.lists
                    .getByTitle(listname)
                    .fields.addNumber(title, options);
                break;
            case "text":
                field = await sp.web.lists
                    .getByTitle(listname)
                    .fields.addText(title, options);
                break;
            case "multi-text":
                field = await sp.web.lists
                    .getByTitle(listname)
                    .fields.addMultilineText(title, options);
                break;
            case "user":
                field = await sp.web.lists
                    .getByTitle(listname)
                    .fields.addUser(title, options);
                if (options.AllowMultipleValues) {
                    await sp.web.lists.getByTitle(listname).fields.getById(field.Id).update({ AllowMultipleValues: true }, "SP.FieldUser");
                }
                break;
            case "boolean":
                field = await sp.web.lists
                    .getByTitle(listname)
                    .fields.addBoolean(title, options);
                break;
            case "datetime":
                field = await sp.web.lists
                    .getByTitle(listname)
                    .fields.addDateTime(title, options);
                break;
            case "currency":
                field = await sp.web.lists
                    .getByTitle(listname)
                    .fields.addCurrency(title, options);
                break;
            case "lookup":
                if (options.LookupListTitle) {
                    const list = await sp.web.lists
                        .getByTitle(options.LookupListTitle)
                        .select("Id")();
                    if (list) {
                        delete options.LookupListTitle;
                        options.LookupListId = list.Id;
                    }
                }
                const allowMultipleValues = options.AllowMultipleValues;
                delete options.AllowMultipleValues;
                field = await sp.web.lists
                    .getByTitle(listname)
                    .fields.addLookup(title, options);
                if (allowMultipleValues) {
                    await sp.web.lists.getByTitle(listname).fields.getById(field.Id).update({ AllowMultipleValues: true }, "SP.FieldLookup");
                }
                break;
            case "choice":
                field = await sp.web.lists
                    .getByTitle(listname)
                    .fields.addChoice(title, options);
                break;
            case "calculated":
                field = await sp.web.lists
                    .getByTitle(listname)
                    .fields.addCalculated(title, {
                        FieldTypeKind: FieldTypes.Calculated,
                        OutputType: FieldTypes.Text,
                        ...options
                    });
                break;
        }
        return field;
    } catch (error) {
        throw new Error(`error while adding field ${title} to list ${listname}: ${error}`);
    }
}

/**
 * Rename an existing Field.
 * @param listname List title
 * @param fieldTitle existing field title
 * @param newTitle new field title
 */
export async function renameField(
    listname: string,
    fieldTitle: string,
    newTitle: string
) {
    try {
        const updateResult = await sp.web.lists
            .getByTitle(listname)
            .fields.getByTitle(fieldTitle)
            .update({
                Title: newTitle,
            });
        return updateResult;
    } catch (error) {
        throw new Error(
            `error while renaming field ${fieldTitle} in list ${listname}: ${error}`
        );
    }
}

/**
 * Add a new View with given Fields to a given List.
 * @param listname List title
 * @param title View title
 * @param personalView is personal View?
 * @param fields Fields to add to the View.
 */
export async function addViewToList(
    listname: string,
    title: string,
    personalView: boolean,
    fields: string[],
    options?: any
) {
    try {
        const view = await sp.web.lists
            .getByTitle(listname)
            .views.add(title, personalView, options);
        // remove view default fields
        await sp.web.lists.getByTitle(listname).views.getById(view.Id).fields.removeAll();
        // add fields to view
        for (const field of fields) {
            await sp.web.lists
                .getByTitle(listname)
                .views.getById(view.Id)
                .fields.add(field);
        }
        // set view query
        if (options.viewXml) {
            await sp.web.lists.getByTitle(listname).views.getById(view.Id).setViewXml(options.viewXml);
        }
        return view;
    } catch (error) {
        throw new Error(`error while adding view to list: ${error}`);
    }
}

/**
 * Returns all Views existing in a List.
 * @param listname List title
 */
export async function getViewsFromList(listname: string): Promise<IViewInfo[]> {
    try {
        const views = await sp.web.lists.getByTitle(listname).views();
        return views;
    } catch (error) {
        throw new Error(`error while getting views from list ${listname}: ${error}`);
    }
}

/**
 * Add a new item to a given List.
 * @param listname List title
 * @param item List item properties
 */
export async function addItemToList(
    listname: string,
    item: any
) {
    try {
        const newItem = await sp.web.lists
            .getByTitle(listname)
            .items.add(item);
        return newItem;
    } catch (error) {
        throw new Error(`error while adding item to list: ${error}`);
    }
}

/**
 * Adds a new file to a document library.
 * @param listname List title
 * @param filePath File Path
 * @param file File Content
 * @param options optional parmeter
 * @param editAssociatedValues new associated item values
 */
export async function addItemToFileCollection(
    listname: string,
    filePath: string,
    file: string | ArrayBuffer | Blob,
    filetype: "base64" | "png",
    options?: any,
    editAssociatedValues?: any
) {
    try {
        let fileContent = file;
        switch (filetype) {
            case "base64":
            case "png":
                const res = await fetch(file as string);
                const blob = await res.blob();
                fileContent = blob;
                break;
        }
        const _file = await sp.web.lists
            .getByTitle(listname)
            .rootFolder.files.addUsingPath(filePath, fileContent, options);
        if (editAssociatedValues) {
            await (await sp.web.lists.getByTitle(listname).rootFolder.files.getByUrl(_file.ServerRelativeUrl).getItem()).update(editAssociatedValues);
        }
        return _file;
    } catch (error) {
        throw new Error(`error while adding item to file collection: ${error}`);
    }
}
/**
 * Deletes a List by its title.
 * @param title List title
 */
export async function deleteList(title: string): Promise<void> {
    try {
        await sp.web.lists.getByTitle(title).recycle();
    } catch (error) {
        throw new Error(`error while deleting list ${title}: ${error}`);
    }
}

/**
 * Returns all Lists contained in the current Web site.
 * @returns Array of Lists contained in the Web site.
 */
export async function getLists(): Promise<IListInfo[]> {
    try {
        const lists: IListInfo[] = await sp.web.lists.select("Title")();
        return lists;
    } catch (error) {
        throw new Error(`error while getting lists from web: ${error}`);
    }
}

export async function getCurrentTimezone() {
    try {
        const timezone = await sp.web.regionalSettings.timeZone();
        return timezone;
    } catch (error) {
        throw new Error(`error while getting current timezone: ${error}`);
    }
}

export async function localTimeToUTC(time: Date): Promise<string> {
    try {
        const utcTime = await sp.web.regionalSettings.timeZone.localTimeToUTC(time);
        return utcTime;
    } catch (error) {
        throw new Error(`error while converting local time to utc: ${error}`);
    }
}

export async function utcTimeToLocal(time: Date): Promise<string> {
    try {
        if (time === null) return null;
        if (time.getFullYear() <= 1970) return null;
        const localTime = await sp.web.regionalSettings.timeZone.utcToLocalTime(time);
        return localTime;
    } catch (error) {
        throw new Error(`error while converting utc time to local: ${error}`);
    }
}

export async function getUserDirectReports(loginName: string) {
    try {
        const userProperties = await getUserProperties(loginName);
        const directReports = [];
        if (userProperties) {
            const directReportLoginNames = userProperties.DirectReports
            if (directReportLoginNames) {
                for (const loginName of directReportLoginNames) {
                    const user = await getUserByLoginName(loginName);
                    directReports.push(user);
                }
            }
            return directReports;
        }
    } catch (error) {
        throw new Error(`error while getting user direct reports: ${error}`);
    }
}

export async function getCurrentUserDirectReports(loginName: string) {
    try {
        const profile = await sp.profiles.myProperties();
        return profile.DirectReports;
    } catch (error) {
        throw new Error(`error while getting user direct reports: ${error}`);
    }
}

export async function getUserManager(loginName: string) {
    try {
        const userProperties = await getUserProperties(loginName);
        if (userProperties) {
            const manager = userProperties.UserProfileProperties?.find(
                (v: any) => v.Key === 'Manager'
            )?.Value;
            if (manager) {
                const managerUser = await getUserByLoginName(manager);
                return managerUser;
            }
        }
    } catch (error) {
        throw new Error(`error while getting user manager: ${error}`);
    }

}

export async function getUserProperties(loginName: string) {
    try {
        const profile = await sp.profiles.getPropertiesFor(loginName);
        return profile;
    } catch (error) {
        throw new Error(`error while getting user properties: ${error}`);
    }
}

export async function getUserByLoginName(loginName: string) {
    try {
        const user = await sp.web.siteUsers.getByLoginName(loginName)();
        return user;
    } catch (error) {
        throw new Error(`error while getting user by login name: ${error}`);
    }
}

export async function getUserByEmail(email: string) {
    try {
        const user = await sp.web.siteUsers.getByEmail(email)();
        return user;
    } catch (error) {
        throw new Error(`error while getting user by email: ${error}`);
    }
}

export async function getUserById(id: number) {
    try {
        const user = await sp.web.siteUsers.getById(id)();
        return user;
    } catch (error) {
        throw new Error(`error while getting user by id: ${error}`);
    }
}

export async function searchUser(searchString: string) {
    try {
        const users = await sp.profiles.clientPeoplePickerSearchUser({
            AllowEmailAddresses: true,
            AllowMultipleEntities: false,
            MaximumEntitySuggestions: 50,
            QueryString: searchString,
        })
        return users;
    } catch (error) {
        throw new Error(`error while searching for user: ${error}`);
    }
}