import {TokenDefinition} from "./tokendefinition";
import {TokensUtils} from "./tokensutils";

export class FieldTitleToken extends TokenDefinition {
    private _value: string = null;

    constructor(web: SP.Web, internalName: string, title: string) {
        super(web, ["{{fieldtitle:" + TokensUtils.escapeRegExp(internalName) + "}}"]);
        this._value = title;
    }

    public getReplaceValue(): Promise<string> {
        return new Promise<string>((resolve, reject) => {
            if (this.cacheValue != null && this.cacheValue !== "") {
                this.cacheValue = this._value;
            }
            resolve(this.cacheValue);
        });
    }
}
