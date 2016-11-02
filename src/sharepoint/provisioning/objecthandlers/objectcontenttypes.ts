"use strict";

import { ObjectHandlerBase } from "./objecthandlerbase";
import { IContentType } from "./../schema/IContentType";
import { IContentTypeFieldRef } from "./../schema/IContentTypeFieldRef";

/**
 * Describes the Object Content Type Handler
 */
export class ObjectContentTypes extends ObjectHandlerBase {
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
    public ProvisionObjects(objects: Array<IContentType>): Promise<{}> {
        super.scope_started();
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
                            this.AddNewContentType(web, contentTypes, fields, currentContentType);
                        } else {
                            // update mode
                            if (currentContentType.Overwrite) {
                                // re-create it
                                existingContentType.deleteObject();
                                clientContext.executeQueryAsync(
                                    () => {
                                        resolve(this.AddNewContentType(web, contentTypes, fields, currentContentType));
                                    });
                            } else {
                                resolve(this.UpdateContentType(web, existingContentType, currentContentType).then(() => { resolve(); }));
                            }
                        }
                    });
                }, (sender, error) => {
                    console.log(error);
                    super.scope_ended();
                    resolve();
                });
        });
    }

    private UpdateContentType(web: SP.Web, existingContentType: SP.ContentType, templateContentType: IContentType): Promise<void> {
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

        return new Promise<void>((resolve, reject) => {
            if (isDirty) {
                existingContentType.update(true);
                web.get_context().executeQueryAsync(() => {
                    // Delta handling
                    let targetIds: SP.Guid[] = existingContentType.get_fieldLinks().get_data().map((item) => { return item.get_id(); });
                    let sourceIds: SP.Guid[] = templateContentType.FieldRefs.map((item) => { return item.ID; });

                    let fieldsNotPresentInTarget = sourceIds.filter((element) => { return targetIds.indexOf(element) < 0; });

                    if (fieldsNotPresentInTarget.length > 0) {
                        fieldsNotPresentInTarget.forEach((fieldId, idx) => {
                            // var fieldRef = templateContentType.FieldRefs.filter(fr => fr.ID === fieldId);
                            let field = web.get_fields().getById(fieldId);
                            existingContentType.get_fields().add(field);
                            // TODO: check if this works...without required and Hidden?
                            // web.AddFieldToContentType(existingContentType, field, fieldRef.Required, fieldRef.Hidden);
                        });
                    }

                    isDirty = false;
                    let commonIds = sourceIds.filter((element) => { return targetIds.indexOf(element) > -1; });
                    commonIds.forEach((fieldId) => {
                        let fieldLink: SP.FieldLink = existingContentType.get_fieldLinks().get_data().filter((fl) => fl.get_id() === fieldId)[0];
                        let fieldRef: IContentTypeFieldRef = templateContentType.FieldRefs.filter(fr => fr.ID === fieldId)[0];
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

                    if (isDirty) {
                        existingContentType.update(true);
                        web.get_context().executeQueryAsync(() => { resolve(); }, () => { reject(); });
                    }
                }, (sender, error) => {
                    console.log(error);
                    reject();
                });
            } else {
                resolve();
            }
        });

    }

    private AddNewContentType(web: SP.Web,
        contentTypes: SP.ContentTypeCollection,
        fields: SP.FieldCollection,
        obj: IContentType): Promise<SP.ContentType> {
        return this.CreateContentType(web, obj, contentTypes.get_data(), fields.get_data());
    }

    private CreateContentType(web: SP.Web,
        templateContentType: IContentType,
        existingCTs: SP.ContentType[] = null,
        existingFields: SP.Field[] = null): Promise<SP.ContentType> {
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
        web.get_context().load(createdCT);
        return new Promise<SP.ContentType>((resolve, reject) => {
            web.get_context().executeQueryAsync(() => {
                if (templateContentType.FieldRefs && templateContentType.FieldRefs.length > 0) {
                    templateContentType.FieldRefs.forEach((fieldRef) => {
                        let field = web.get_fields().getById(fieldRef.ID);
                        let fieldLink: SP.FieldLinkCreationInformation = new SP.FieldLinkCreationInformation();
                        fieldLink.set_field(field);
                        createdCT.get_fieldLinks().add(fieldLink);
                        // web.addFieldToContentType(createdCT, field, fieldRef.Required, fieldRef.Hidden);
                    });
                    createdCT.get_fieldLinks().reorder(templateContentType.FieldRefs.map((field) => { return field.Name; }));
                }

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
                    // TODO: After finding the proper API for Document set to update this part
                    resolve(createdCT);
                }, (sender, error) => {
                    console.log(error);
                    reject(null);
                });
            }, (sender, error) => {
                console.log(error);
                reject(null);
            });

        });
    }
}
