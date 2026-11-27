/* eslint-disable @typescript-eslint/no-explicit-any */
import { ISPFXContext, SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/attachments";
import "@pnp/sp/profiles";
import "@pnp/sp/site-users";
import "@pnp/sp/fields";
import "@pnp/sp/views";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/regional-settings/web";
import { IListEnsureResult, IListInfo } from "@pnp/sp/lists/types";
import { FieldTypes, IFieldInfo } from "@pnp/sp/fields";
import { IViewInfo } from "@pnp/sp/views/types";

let sp: SPFI;

/**
 * Initializes the PnP SharePoint Object. Must be called once in the web part's onInit.
 * @param context SP page context
 */
export function initSP(context: ISPFXContext): void {
    sp = spfi().using(SPFx(context));
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
 * Gets all List Items from a List. Pages through the list in chunks of 5000 items.
 * @param listTitle List Name
 * @param withAttachments adds the attachment files as `Attachments` to each item
 * @param queryFilter raw OData filter, escape user input (e.g. `'` → `''`) before passing it
 * @param fields fields to select
 * @param lookup lookup fields to expand
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
        let query = sp.web.lists.getByTitle(listTitle).items;
        const expand = [...(lookup ?? []), ...(withAttachments ? ["AttachmentFiles"] : [])];

        if (fields) {
            query = query.select(...fields, ...(withAttachments ? ["AttachmentFiles"] : []));
        }
        if (expand.length) {
            query = query.expand(...expand);
        }
        if (queryFilter) {
            query = query.filter(queryFilter);
        }

        const items: any[] = [];
        for await (const page of query.top(5000)) {
            items.push(...page);
        }

        if (withAttachments) {
            for (const item of items) {
                item.Attachments = item.AttachmentFiles;
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
export async function createListItem(listTitle: string, item: Record<string, any>): Promise<any> {
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
export async function addAttachmentToListItem(listTitle: string, itemID: number, file: Blob | ArrayBuffer | string, fileName: string): Promise<any> {
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
export async function recycleAttachment(listTitle: string, itemID: number, fileName: string): Promise<void> {
    try {
        await sp.web.lists
            .getByTitle(listTitle)
            .items.getById(itemID)
            .attachmentFiles.getByName(fileName)
            .recycle();
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
export async function getListItemById(listTitle: string, itemID: number): Promise<any> {
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
    updateValue: Record<string, any>
): Promise<any> {
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

export async function deleteListItemById(listTitle: string, itemID: number): Promise<void> {
    try {
        await sp.web.lists
            .getByTitle(listTitle)
            .items.getById(itemID)
            .delete();
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
export async function getCurrentUserProfile(): Promise<any> {
    try {
        const profile = await sp.web.currentUser();
        return profile;
    } catch (error) {
        throw new Error(`error while getting current user profile ${error}`);
    }
}

export async function getListFields(listTitle: string): Promise<IFieldInfo[]> {
    try {
        const fields = await sp.web.lists.getByTitle(listTitle).fields();
        return fields;
    } catch (error) {
        throw new Error(`error while getting fields from list ${listTitle}: ${error}`);
    }
}

/**
 * Add a new List.
 * @param title List title
 * @param description List description
 * @param template List template id
 */
export async function addList(
    title: string,
    description?: string,
    template?: number
): Promise<IListInfo> {
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
 * @param description List description
 * @param template List template id
 * @param additionalSettings additional list properties
 */
export async function ensureList(
    title: string,
    description?: string,
    template?: number,
    additionalSettings?: Partial<IListInfo>
): Promise<IListEnsureResult> {
    try {
        const list = await sp.web.lists.ensure(title, description, template, false, additionalSettings);
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
    type: "text" | "multi-text" | "number" | "user" | "boolean" | "datetime" | "currency" | "lookup" | "choice" | "calculated",
    options: Record<string, any> = {}
): Promise<Partial<IFieldInfo> | undefined> {
    try {
        const listFields = sp.web.lists.getByTitle(listname).fields;
        const fieldOptions = { ...options };
        let field: Partial<IFieldInfo> | undefined;
        switch (type) {
            case "number":
                field = await listFields.addNumber(title, fieldOptions);
                break;
            case "text":
                field = await listFields.addText(title, fieldOptions);
                break;
            case "multi-text":
                field = await listFields.addMultilineText(title, fieldOptions);
                break;
            case "user":
                field = await listFields.addUser(title, fieldOptions as any);
                if (fieldOptions.AllowMultipleValues && field.Id) {
                    await listFields.getById(field.Id).update({ AllowMultipleValues: true }, "SP.FieldUser");
                }
                break;
            case "boolean":
                field = await listFields.addBoolean(title, fieldOptions);
                break;
            case "datetime":
                field = await listFields.addDateTime(title, fieldOptions);
                break;
            case "currency":
                field = await listFields.addCurrency(title, fieldOptions);
                break;
            case "lookup": {
                if (fieldOptions.LookupListTitle) {
                    const list = await sp.web.lists
                        .getByTitle(fieldOptions.LookupListTitle)
                        .select("Id")();
                    delete fieldOptions.LookupListTitle;
                    fieldOptions.LookupListId = list.Id;
                }
                const allowMultipleValues = fieldOptions.AllowMultipleValues;
                delete fieldOptions.AllowMultipleValues;
                field = await listFields.addLookup(title, fieldOptions as any);
                if (allowMultipleValues && field.Id) {
                    await listFields.getById(field.Id).update({ AllowMultipleValues: true }, "SP.FieldLookup");
                }
                break;
            }
            case "choice":
                field = await listFields.addChoice(title, fieldOptions as any);
                break;
            case "calculated":
                field = await listFields.addCalculated(title, {
                    FieldTypeKind: FieldTypes.Calculated,
                    OutputType: FieldTypes.Text,
                    ...fieldOptions
                } as any);
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
): Promise<void> {
    try {
        await sp.web.lists
            .getByTitle(listname)
            .fields.getByTitle(fieldTitle)
            .update({
                Title: newTitle,
            });
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
 * @param options additional view properties, `viewXml` is applied after creation
 */
export async function addViewToList(
    listname: string,
    title: string,
    personalView: boolean,
    fields: string[],
    options: Record<string, any> & { viewXml?: string } = {}
): Promise<IViewInfo> {
    try {
        const { viewXml, ...additionalSettings } = options;
        const views = sp.web.lists.getByTitle(listname).views;
        const view = await views.add(title, personalView, additionalSettings);
        const viewFields = views.getById(view.Id).fields;
        await viewFields.removeAll();
        for (const field of fields) {
            await viewFields.add(field);
        }
        if (viewXml) {
            await views.getById(view.Id).setViewXml(viewXml);
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
    item: Record<string, any>
): Promise<any> {
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
 * @param file File Content, for "base64"/"png" a data URL
 * @param filetype "base64"/"png" converts the data URL to a Blob
 * @param options optional parameter
 * @param editAssociatedValues new associated item values
 */
export async function addItemToFileCollection(
    listname: string,
    filePath: string,
    file: string | ArrayBuffer | Blob,
    filetype?: "base64" | "png",
    options?: Record<string, any>,
    editAssociatedValues?: Record<string, any>
): Promise<any> {
    try {
        let fileContent = file;
        if (filetype === "base64" || filetype === "png") {
            const res = await fetch(file as string);
            fileContent = await res.blob();
        }
        const files = sp.web.lists.getByTitle(listname).rootFolder.files;
        const addedFile = await files.addUsingPath(filePath, fileContent, options);
        if (editAssociatedValues) {
            const item = await files.getByUrl(addedFile.ServerRelativeUrl).getItem();
            await item.update(editAssociatedValues);
        }
        return addedFile;
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

export async function getCurrentTimezone(): Promise<any> {
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

// eslint-disable-next-line @rushstack/no-new-null
export async function utcTimeToLocal(time: Date | null): Promise<string | null> {
    try {
        if (time === null) return null;
        if (time.getFullYear() <= 1970) return null;
        const localTime = await sp.web.regionalSettings.timeZone.utcToLocalTime(time);
        return localTime;
    } catch (error) {
        throw new Error(`error while converting utc time to local: ${error}`);
    }
}

export async function getUserDirectReports(loginName: string): Promise<any[]> {
    try {
        const userProperties = await getUserProperties(loginName);
        const directReports = [];
        for (const directReportLoginName of userProperties?.DirectReports ?? []) {
            directReports.push(await getUserByLoginName(directReportLoginName));
        }
        return directReports;
    } catch (error) {
        throw new Error(`error while getting user direct reports: ${error}`);
    }
}

export async function getCurrentUserDirectReports(): Promise<string[]> {
    try {
        const profile = await sp.profiles.myProperties();
        return profile.DirectReports;
    } catch (error) {
        throw new Error(`error while getting user direct reports: ${error}`);
    }
}

export async function getUserManager(loginName: string): Promise<any> {
    try {
        const userProperties = await getUserProperties(loginName);
        const manager = userProperties?.UserProfileProperties?.find(
            (v: any) => v.Key === 'Manager'
        )?.Value;
        if (manager) {
            return await getUserByLoginName(manager);
        }
        return undefined;
    } catch (error) {
        throw new Error(`error while getting user manager: ${error}`);
    }
}

export async function getUserProperties(loginName: string): Promise<any> {
    try {
        const profile = await sp.profiles.getPropertiesFor(loginName);
        return profile;
    } catch (error) {
        throw new Error(`error while getting user properties: ${error}`);
    }
}

export async function getUserByLoginName(loginName: string): Promise<any> {
    try {
        const user = await sp.web.siteUsers.getByLoginName(loginName)();
        return user;
    } catch (error) {
        throw new Error(`error while getting user by login name: ${error}`);
    }
}

export async function getUserByEmail(email: string): Promise<any> {
    try {
        const user = await sp.web.siteUsers.getByEmail(email)();
        return user;
    } catch (error) {
        throw new Error(`error while getting user by email: ${error}`);
    }
}

export async function getUserById(id: number): Promise<any> {
    try {
        const user = await sp.web.siteUsers.getById(id)();
        return user;
    } catch (error) {
        throw new Error(`error while getting user by id: ${error}`);
    }
}

export async function searchUser(searchString: string): Promise<any> {
    try {
        const users = await sp.profiles.clientPeoplePickerSearchUser({
            AllowEmailAddresses: true,
            AllowMultipleEntities: false,
            MaximumEntitySuggestions: 50,
            QueryString: searchString,
        });
        return users;
    } catch (error) {
        throw new Error(`error while searching for user: ${error}`);
    }
}
