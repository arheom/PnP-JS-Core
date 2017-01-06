import {TokenDefinition} from "./tokendefinition";
import {TokensUtils} from "./tokensutils";

export class ParameterToken extends TokenDefinition {
    private _value: string = null;

    constructor(web: SP.Web, name: string, value: string) {
        super(web, ["<<parameter:" + TokensUtils.escapeRegExp(name) + ">>", "<<\\$" + TokensUtils.escapeRegExp(name) + ">>"]);
        this._value = value.toString();
    }

    public getReplaceValue(): Promise<string> {
        return new Promise<string>((resolve, reject) => {
            if (this.cacheValue == null || this.cacheValue === "") {
                this.cacheValue = this._value;
            }
            resolve(this.cacheValue);
        });
    }
}
