import {TokenDefinition} from "./tokendefinition";
import {TokensUtils} from "./tokensutils";

export class TermStoreIdToken extends TokenDefinition {
    private _value: string = null;

    constructor(web: SP.Web, storeName: string, id: SP.Guid) {
        super(web, ["<<termstoreid:" + TokensUtils.escapeRegExp(storeName) + ">>"]);
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
