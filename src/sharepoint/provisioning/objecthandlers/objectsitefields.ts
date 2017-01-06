"use strict";

import { ObjectHandlerBase } from "./objecthandlerbase";
import { TokenParser } from "./tokenparser";
import { IField } from "../schema/IField";
import { Logger, LogLevel } from "../../../utils/logging";

/**
 * Describes the SiteFields Object Handler
 */
export class ObjectSiteFields extends ObjectHandlerBase {

    private contentXmlJsonProperty: string = "content";
    /**
     * Creates a new instance of the ObjectLists class
     */
    constructor() {
        super("SiteFields");
    }

    /**
     * Provision Lists
     * 
     * @param objects The lists to provision
     */
    public ProvisionObjects(objects: Array<IField>, tokenParser: TokenParser): Promise<{}> {
        super.scope_started();
        return new Promise<{}>((resolve, reject) => {
            const clientContext = SP.ClientContext.get_current();
            // fields will be provisioned inside the root web, not in subsites
            const web = clientContext.get_site().get_rootWeb();
            let existingFields: SP.FieldCollection = web.get_fields();
            clientContext.load(existingFields, "Include(Id, InternalName)");
            clientContext.executeQueryAsync(
                () => {
                    objects.forEach((currentField) => {
                        let schemaXml = currentField.SchemaXml ? currentField.SchemaXml : this.GetFieldXml(currentField);
                        tokenParser.parseString(schemaXml).then(currentFieldSchemaXml => {
                            let existingField: SP.Field = existingFields.get_data().filter((field) => {
                                try {
                                    return field.get_id().toString().toLowerCase() === currentField.ID.toLowerCase();
                                } catch (error) {
                                    Logger.write(error.toString(), LogLevel.Error);
                                    return false;
                                    // from some reasons, even was specifically loaded, the ids of some fields might not be loaded
                                }
                            })[0];
                            if (existingField) {
                                // update mode
                                this.UpdateField(web, existingField.get_id().toString(), currentFieldSchemaXml).then((field: SP.Field) => {
                                    resolve(field);
                                });
                            } else {
                                // add mode
                                this.CreateField(web, currentFieldSchemaXml).then((field: SP.Field) => {
                                    resolve(field);
                                });
                            }
                        });
                    });
                });
        });
    }


    private UpdateField(web: SP.Web, fieldId: string, originalFieldXml: string): Promise<SP.Field> {
        return new Promise<SP.Field>((resolve, reject) => {
            let existingField = web.get_fields().getById(new SP.Guid(fieldId));
            web.get_context().load(existingField, "SchemaXml");
            web.get_context().executeQueryAsync(() => {
                let schemaXml = existingField.get_schemaXml();
                if (schemaXml.toLocaleLowerCase() !== originalFieldXml.toLocaleLowerCase()) {
                    // needs update
                    let parser = new DOMParser();
                    let existingFieldElement = parser.parseFromString(schemaXml, "text/xml").getElementsByTagName("Field")[0] as HTMLElement;
                    let templateFieldElement = parser.parseFromString(originalFieldXml, "text/xml").getElementsByTagName("Field")[0] as HTMLElement;
                    if (existingFieldElement.getAttribute("Type") === templateFieldElement.getAttribute("Type")) {
                        // Is existing field of the same type?
                        let listIdentifier = templateFieldElement.getAttribute("List") != null ? templateFieldElement.getAttribute("List") : null;

                        if (listIdentifier != null) {
                            // Temporary remove list attribute from list
                            templateFieldElement.removeAttribute("List");
                        }

                        for (let i = 0; i < templateFieldElement.attributes.length; i++) {
                            let attribute = templateFieldElement.attributes[i];
                            if (existingFieldElement.getAttribute(attribute.name) != null) {
                                existingFieldElement.setAttribute(attribute.name, attribute.value);
                            } else {
                                existingFieldElement.attributes.setNamedItem(attribute.cloneNode(false) as Attr);
                            }
                        }

                        for (let i = 0; i < templateFieldElement.childNodes.length; i++) {
                            let element = templateFieldElement.childNodes[i];
                            if (existingFieldElement.getElementsByTagName(element.nodeName).length > 0) {
                                for (let j = 0; j < existingFieldElement.getElementsByTagName(element.nodeName).length; j++) {
                                    let elementToRemove = existingFieldElement.getElementsByTagName(element.nodeName)[j];
                                    existingFieldElement.removeChild(elementToRemove);
                                }
                            }
                            existingFieldElement.appendChild(element);
                        }

                        if (existingFieldElement.getAttribute("Version") != null) {
                            existingFieldElement.removeAttribute("Version");
                        }
                        existingField.set_schemaXml(this.ReplaceBooleanStringToUpper(existingFieldElement.outerHTML));
                        existingField.updateAndPushChanges(true);
                        web.get_context().load(existingField, "TypeAsString", "DefaultValue");
                        web.get_context().executeQueryAsync(() => {
                            if ((existingField.get_typeAsString() === "TaxonomyFieldType" ||
                                existingField.get_typeAsString() === "TaxonomyFieldTypeMulti") &&
                                (existingField.get_defaultValue() != null && existingField.get_defaultValue() !== "")) {
                                let taxField = web.get_context().castTo(existingField, SP.Taxonomy.TaxonomyField) as SP.Taxonomy.TaxonomyField;
                                this.ValidateTaxonomyFieldDefaultValue(taxField).then(() => {
                                    resolve(taxField);
                                }, () => { reject(); });

                            } else {
                                resolve(existingField);
                            }
                        },
                            (sender, error) => {
                                Logger.write(error.toString(), LogLevel.Error);
                                reject(error);
                            });
                    }
                }
            },
                (sender, error) => {
                    Logger.write(error.toString(), LogLevel.Error);
                    reject(error);
                });
        });
    }

    private CreateField(web: SP.Web, originalFieldXml: string): Promise<SP.Field> {
        return new Promise<SP.Field>((resolve, reject) => {
            let parser = new DOMParser();
            let templateFieldElement = parser.parseFromString(originalFieldXml, "text/xml").getElementsByTagName("Field")[0] as HTMLElement;
            let listIdentifier = templateFieldElement.getAttribute("List") != null ? templateFieldElement.getAttribute("List") : null;

            if (listIdentifier != null) {
                // Temporary remove list attribute from list
                templateFieldElement.removeAttribute("List");
            }

            let fieldXml = this.ReplaceBooleanStringToUpper(templateFieldElement.outerHTML);
            let field = web.get_fields().addFieldAsXml(fieldXml, false, SP.AddFieldOptions.addFieldInternalNameHint);
            web.get_context().load(field, "Id", "TypeAsString", "DefaultValue", "InternalName", "Title", "SchemaXml");
            web.get_context().executeQueryAsync(() => {
                if ((field.get_typeAsString() === "TaxonomyFieldType" ||
                    field.get_typeAsString() === "TaxonomyFieldTypeMulti")
                    && (field.get_defaultValue() != null && field.get_defaultValue() !== "")) {
                    let taxField = web.get_context().castTo(field, SP.Taxonomy.TaxonomyField) as SP.Taxonomy.TaxonomyField;
                    this.ValidateTaxonomyFieldDefaultValue(taxField).then(() => {
                        resolve(taxField);
                    });
                } else {
                    resolve(field);
                }
            },
                (sender, error) => {
                    Logger.write(error.toString(), LogLevel.Error);
                    reject(error);
                });
        });
    }

    private GetFieldXml(field: IField) {
        let fieldXml = "";
        if (!field.SchemaXml) {
            let properties = [];
            let children = [];
            let xmlContent = "";
            Object.keys(field).forEach(prop => {
                let value = field[prop];
                if (this.IsPrimitiveValue(value)) {
                    if (prop.toLowerCase() !== this.contentXmlJsonProperty) {
                        properties.push(`${prop}="${value}"`);
                    } else {
                        xmlContent = value;
                    }
                } else {
                    children.push({ key: prop, value: value });
                }
            });
            // Because of JSON - XML transformation, we need one property to act as Content of an XML Node. By convention, this property is named "Content"
            let fieldInnerXml = children.map(element => this.GetRecursiveFieldXml(element)).join(" ");
            if (fieldInnerXml === "" && xmlContent !== "") {
                fieldInnerXml = xmlContent;
            }
            fieldXml = `<Field ${properties.join(" ")}>`;
            if (field.Type === "Calculated") {
                fieldXml += `<Formula>${field.Formula}</Formula>`;
            }
            fieldXml += fieldInnerXml;
            fieldXml += "</Field>";
        }
        return fieldXml;
    }

    private ReplaceBooleanStringToUpper(input: string): string {
        let output = "";
        output = input.replace(new RegExp("\"true\"", "g"), "\"TRUE\"");
        output = output.replace(new RegExp("\"false\"", "g"), "\"FALSE\"");
        return output;
    }

    private GetRecursiveFieldXml(element: { key: any, value: any }): string {
        if (element == null) {
            return "";
        }

        let properties = [];
        let children = [];
        let xmlContent = "";
        if (!isNaN(parseInt(element.key, 0))) {
           // this is part of the array - replace it with property
            element.key = "Property";
        }
        let elementXml = "<" + element.key + " {attributes}>{children}</" + element.key + ">";
        Object.keys(element.value).forEach(prop => {
            let value = element.value[prop];
            if (this.IsPrimitiveValue(value)) {
                if (prop.toLowerCase() !== this.contentXmlJsonProperty) {
                    properties.push(`${prop}="${value}"`);
                } else {
                    xmlContent = value;
                }
            } else {
                children.push({ key: prop, value: value });
            }
        });
        elementXml = elementXml.replace("{attributes}", properties.join(" "));
        let fieldInnerXml = "";
        if (children.length > 0) {
            fieldInnerXml = children.map(childElement => this.GetRecursiveFieldXml(childElement)).join(" ");
        } else {
            fieldInnerXml = xmlContent;
        }
        elementXml = elementXml.replace("{children}", fieldInnerXml);
        return elementXml;
    }

    private IsPrimitiveValue(value: any) {
        return (typeof (value) === "boolean" ||
            typeof (value) === "number" ||
            typeof (value) === "string");
    }

    private ValidateTaxonomyFieldDefaultValue(field: SP.Taxonomy.TaxonomyField): Promise<void> {
        // get validated value with correct WssIds
        return new Promise<void>((resolve, reject) => {
            this.GetTaxonomyFieldValidatedValue(field.get_context() as SP.ClientContext, field, field.get_defaultValue()).then(validatedValue => {
                if ((validatedValue != null && validatedValue !== "") && field.get_defaultValue() !== validatedValue) {
                    field.set_defaultValue(validatedValue);
                    field.updateAndPushChanges(true);
                    field.get_context().executeQueryAsync(() => { resolve(); }, (sender, error) => { reject(error.get_message()); });
                } else {
                    resolve();
                }
            });
        });
    }

    private GetTaxonomyFieldValidatedValue(context: SP.ClientContext,
        field: SP.Taxonomy.TaxonomyField,
        defaultValue: string): Promise<string> {
        let res: string = null;
        let parsedValue: any = null;
        return new Promise<string>((resolve, reject) => {
            if (field.get_allowMultipleValues()) {
                parsedValue = new SP.Taxonomy.TaxonomyFieldValueCollection(context, defaultValue, field);
            } else {
                let taxValue: SP.Taxonomy.TaxonomyFieldValue = this.TryParseTaxonomyFieldValue(defaultValue);
                if (taxValue) {
                    parsedValue = taxValue;
                }
            }
            if (parsedValue != null) {
                let validateValue = field.getValidatedString(parsedValue);
                context.executeQueryAsync(() => {
                    resolve(validateValue.get_value());
                },
                    () => { reject(); }
                );
            } else {
                resolve(res);
            }
        });
    }

    private TryParseTaxonomyFieldValue(value: string): SP.Taxonomy.TaxonomyFieldValue {
        let taxValue: SP.Taxonomy.TaxonomyFieldValue = null;
        let res = false;
        taxValue = new SP.Taxonomy.TaxonomyFieldValue();
        if (value != null && value !== "") {
            let split: string[] = value.split(";#");
            let wssId = split[0];

            if (split.length > 0 && wssId != null && wssId !== "") {
                taxValue.set_wssId(parseInt(wssId, 0));
                res = true;
            }

            if (res && split.length === 2) {
                let term = split[1];
                let splitTerm = term.split("|");
                let termId = SP.Guid.get_empty();
                if (splitTerm.length > 0) {
                    termId = new SP.Guid(splitTerm[splitTerm.length - 1]);
                    res = termId.toString() !== SP.Guid.get_empty().toString();
                    taxValue.set_termGuid(termId);
                    if (res && splitTerm.length > 1) {
                        taxValue.set_label(splitTerm[0]);
                    }
                } else {
                    res = false;
                }
                res = true;
            } else {

                if (split.length === 1 && wssId != null && wssId !== "") {
                    taxValue.set_wssId(parseInt(wssId, 0));
                    res = true;
                }
            }
        }
        if (res) {
            return taxValue;
        } else {
            return null;
        }
    }
}
