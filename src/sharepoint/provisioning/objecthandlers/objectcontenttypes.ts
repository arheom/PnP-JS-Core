"use strict";

import { ObjectHandlerBase } from "./objecthandlerbase";
import { IContentType } from "./../schema/IContentType";
import { IContentTypeFieldRef } from "./../schema/IContentTypeFieldRef";
import { Logger, LogLevel } from "../../../utils/logging";
import { TokenParser } from "./tokenparser";
import { ContentTypeIdToken } from "./tokendefinitions/contenttypeidtoken";

/**
 * Describes the Object Content Type Handler
 */
export class ObjectContentTypes extends ObjectHandlerBase {
    private tokenParser: TokenParser;
    /**
     * Creates a new instance of the ObjectContentType class
     */
    constructor() {
        super("ContentType");
    }

    /**
     * Provisioning Content Type
     * 
     * @param customactions The Content Types to provision
     */
    public ProvisionObjects(objects: Array<IContentType>, tokenParser: TokenParser): Promise<{}> {
        super.scope_started();
        this.tokenParser = tokenParser;
        return new Promise<{}>((resolve, reject) => {
            const clientContext = SP.ClientContext.get_current();
            // content types will be provisioned inside the root web, not in subsites
            const web = clientContext.get_site().get_rootWeb();
            let contentTypes: SP.ContentTypeCollection = web.get_contentTypes();
            let fields: SP.FieldCollection = web.get_fields();
            clientContext.load(contentTypes, "Include(FieldLinks, Name, ReadOnly, Sealed, Hidden, Description, " +
                "DocumentTemplate, Group, DisplayFormUrl, EditFormUrl, NewFormUrl, FieldLinks.Include(Required, Hidden))");
            clientContext.load(fields);
            clientContext.executeQueryAsync(
                () => {
                    objects.forEach((currentContentType) => {
                        let existingContentType: SP.ContentType = contentTypes.get_data().filter((contentType) => {
                            return contentType.get_name() === currentContentType.Name;
                        })[0];

                        if (existingContentType == null) {
                            // add mode
                            resolve(this.CreateContentType(web, currentContentType, contentTypes.get_data(), fields.get_data()));
                        } else {
                            // update mode
                            if (currentContentType.Overwrite) {
                                // re-create it
                                existingContentType.deleteObject();
                                clientContext.executeQueryAsync(
                                    () => {
                                        resolve(this.CreateContentType(web, currentContentType, contentTypes.get_data(), fields.get_data()));
                                    });
                            } else {
                                resolve(this.UpdateContentType(web, existingContentType, currentContentType).then(() => { resolve(); }));
                            }
                        }
                    });
                }, (sender, error) => {
                    Logger.write(error.toString(), LogLevel.Error);
                    super.scope_ended();
                    resolve();
                });
        });
    }

    private UpdateContentType(web: SP.Web, existingContentType: SP.ContentType, templateContentType: IContentType): Promise<boolean> {
        let isDirty = false;
        if (templateContentType.Hidden && existingContentType.get_hidden() !== templateContentType.Hidden) {
            existingContentType.set_hidden(templateContentType.Hidden);
            isDirty = true;
        }
        if (templateContentType.ReadOnly && existingContentType.get_readOnly() !== templateContentType.ReadOnly) {
            existingContentType.set_readOnly(templateContentType.ReadOnly);
            isDirty = true;
        }
        if (templateContentType.Sealed && existingContentType.get_sealed() !== templateContentType.Sealed) {
            existingContentType.set_sealed(templateContentType.Sealed);
            isDirty = true;
        }
        if (templateContentType.Description && existingContentType.get_description() !== templateContentType.Description) {
            existingContentType.set_description(templateContentType.Description);
            isDirty = true;
        }
        if (templateContentType.DocumentTemplate &&
            existingContentType.get_documentTemplate() !== templateContentType.DocumentTemplate) {
            existingContentType.set_documentTemplate(templateContentType.DocumentTemplate);
            isDirty = true;
        }
        if (templateContentType.Name && existingContentType.get_name() !== templateContentType.Name) {
            existingContentType.set_name(templateContentType.Name);
            isDirty = true;
        }
        if (templateContentType.Group && existingContentType.get_group() !== templateContentType.Group) {
            existingContentType.set_group(templateContentType.Group);
            isDirty = true;
        }
        if (templateContentType.DisplayFormUrl && existingContentType.get_displayFormUrl() !== templateContentType.DisplayFormUrl) {
            existingContentType.set_displayFormUrl(templateContentType.DisplayFormUrl);
            isDirty = true;
        }
        if (templateContentType.EditFormUrl && existingContentType.get_editFormUrl() !== templateContentType.EditFormUrl) {
            existingContentType.set_editFormUrl(templateContentType.EditFormUrl);
            isDirty = true;
        }
        if (templateContentType.NewFormUrl && existingContentType.get_newFormUrl() !== templateContentType.NewFormUrl) {
            existingContentType.set_newFormUrl(templateContentType.NewFormUrl);
            isDirty = true;
        }
        // TODO: Check how to get the info for SP2013
        // #if !SP2013
        //   if (templateContentType.Name.ContainsResourceToken()) {
        //  existingContentType.NameResource.SetUserResourceValue(templateContentType.Name, parser);
        //  isDirty = true;
        // }
        // if (templateContentType.Description.ContainsResourceToken()) {
        // existingContentType.DescriptionResource.SetUserResourceValue(templateContentType.Description, parser);
        // isDirty = true;
        // }
        // #endif

        return new Promise<boolean>((resolve, reject) => {
            if (isDirty) {
                existingContentType.update(true);
                web.get_context().load(existingContentType);
                web.get_context().executeQueryAsync(() => {
                    resolve(this.UpdateContentTypeFields(web, existingContentType, templateContentType));
                }, (sender, error) => {
                    Logger.write(error.toString(), LogLevel.Error);
                    reject();
                });
            } else {
                resolve(this.UpdateContentTypeFields(web, existingContentType, templateContentType));
            }
        });

    }

    private UpdateContentTypeFields(web: SP.Web, existingContentType: SP.ContentType, templateContentType: IContentType): Promise<boolean> {
        return new Promise<boolean>((resolve, reject) => {
            let isDirty = false;
            // Delta handling
            let targetIds: SP.Guid[] = existingContentType.get_fieldLinks().get_data().map((item) => item.get_id());
            let targetAllIds:string = targetIds.map(id => id.toString()).join(",").toLowerCase();
            let sourceIds: SP.Guid[] = templateContentType.FieldRefs.map((item) => item.ID);
            let commonIds = sourceIds.filter((element) => targetAllIds.indexOf(element.toString().toLowerCase()) > -1);
            let fieldsNotPresentInTarget = sourceIds.filter((element) => targetAllIds.indexOf(element.toString().toLowerCase()) < 0);

            commonIds.forEach((fieldId) => {
                let fieldLink: SP.FieldLink = existingContentType.get_fieldLinks().get_data().filter((fl) => fl.get_id().toString().toLowerCase() === fieldId.toString().toLowerCase())[0];
                let fieldRef: IContentTypeFieldRef = templateContentType.FieldRefs.filter(fr => fr.ID.toString().toLowerCase() === fieldId.toString().toLowerCase())[0];
                if (fieldRef != null) {
                    if (fieldLink.get_required() !== fieldRef.Required) {
                        fieldLink.set_required(fieldRef.Required);
                        isDirty = true;
                    }
                    if (fieldLink.get_hidden() !== fieldRef.Hidden) {
                        fieldLink.set_hidden(fieldRef.Hidden);
                        isDirty = true;
                    }
                }
            });

            if (fieldsNotPresentInTarget.length > 0) {
                fieldsNotPresentInTarget.forEach((fieldId: SP.Guid) => {
                    let fieldLink = new SP.FieldLinkCreationInformation();
                    fieldLink.set_field(web.get_fields().getById(fieldId));
                    existingContentType.get_fieldLinks().add(fieldLink);
                });
                existingContentType.get_fieldLinks().reorder(templateContentType.FieldRefs.map((field) => { return field.Name; }));
                existingContentType.update(true);
                web.get_context().executeQueryAsync(() => {
                    resolve(true);
                }, (sender, error) => {
                    Logger.write(error.toString(), LogLevel.Error);
                    reject();
                });
            } else {
                if (isDirty) {
                    existingContentType.update(true);
                    web.get_context().executeQueryAsync(() => {
                        resolve(true);
                    }, (sender, error) => {
                        Logger.write(error.toString(), LogLevel.Error);
                        reject();
                    });
                } else {
                    resolve(true);
                }
            }
        });
    }

    private CreateContentType(web: SP.Web,
            templateContentType: IContentType,
            existingCTs: SP.ContentType[] = null,
            existingFields: SP.Field[] = null): Promise<boolean> {
        let name = templateContentType.Name;
        let description = templateContentType.Description;
        let parentId = templateContentType.ParentId.toString();
        let group = templateContentType.Group;

        let createdCTInfo = new SP.ContentTypeCreationInformation();
        createdCTInfo.set_name(name);
        if (description) {
            createdCTInfo.set_description(description);
        }
        if (group) {
            createdCTInfo.set_group(group);
        }
        // TODO: not sure how to set the ID of the content type and not only the parent ID
        createdCTInfo.set_parentContentType(web.get_contentTypes().getById(parentId));
        let createdCT: SP.ContentType = web.get_contentTypes().add(createdCTInfo);
        return new Promise<boolean>((resolve, reject) => {
            web.get_context().load(createdCT, "Id", "Name");
            web.get_context().executeQueryAsync(() => {
                // add the token to be used later in the installation
                this.tokenParser.addToken(new ContentTypeIdToken(web, templateContentType.Name, createdCT.get_id().toString()));

                if (templateContentType.ReadOnly) {
                    createdCT.set_readOnly(templateContentType.ReadOnly);
                }
                if (templateContentType.Hidden) {
                    createdCT.set_hidden(templateContentType.Hidden);
                }
                if (templateContentType.Sealed) {
                    createdCT.set_sealed(templateContentType.Sealed);
                }

                if (templateContentType.DocumentSetTemplate) {
                    // Only apply a document template when the contenttype is not a document set
                    if (templateContentType.DocumentTemplate !== "") {
                        createdCT.set_documentTemplate(templateContentType.DocumentTemplate);
                    }
                }

                if (templateContentType.NewFormUrl && templateContentType.NewFormUrl !== "") {
                    createdCT.set_newFormUrl(templateContentType.NewFormUrl);
                }
                if (templateContentType.EditFormUrl && templateContentType.EditFormUrl !== "") {
                    createdCT.set_editFormUrl(templateContentType.EditFormUrl);
                }
                if (templateContentType.DisplayFormUrl && templateContentType.DisplayFormUrl !== "") {
                    createdCT.set_displayFormUrl(templateContentType.DisplayFormUrl);
                }

                createdCT.update(true);
                web.get_context().load(createdCT);
                web.get_context().executeQueryAsync(() => {
                    resolve(this.UpdateContentTypeFields(web, createdCT, templateContentType));
                }, (sender, error) => {
                    Logger.write(error.toString(), LogLevel.Error);
                    reject();
                });
            }, (sender, error) => {
                Logger.write(error.toString(), LogLevel.Error);
                reject();
                });
        });
    }
}
