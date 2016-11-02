import {TokenDefinition} from "./tokendefinition";

export class RoleDefinitionToken extends TokenDefinition {
    private name: string = null;

    constructor(web: SP.Web, definition: SP.RoleDefinition) {
        super(web, ["{{roledefinition:" + definition.get_roleTypeKind().toString() + "}}"]);
        this.name = definition.get_name();
    }

    public getReplaceValue(): Promise<string> {
        return new Promise<string>((resolve, reject) => {
            if (this.cacheValue != null && this.cacheValue !== "") {
                this.cacheValue = this.name;
            }
            resolve(this.cacheValue);
        });
    }
}
