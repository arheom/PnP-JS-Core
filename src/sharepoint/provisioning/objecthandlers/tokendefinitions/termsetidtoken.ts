import {TokenDefinition} from "./tokendefinition";
import {TokensUtils} from "./tokensutils";

export class TermSetIdToken extends TokenDefinition {
    private _value: string = null;

    constructor(web: SP.Web, groupName: string, termsetName: string, id: SP.Guid) {
        super(web, ["<<termsetid:" + TokensUtils.escapeRegExp(groupName) + ":" + TokensUtils.escapeRegExp(termsetName) + ">>"]);
        this._value = id.toString();
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
