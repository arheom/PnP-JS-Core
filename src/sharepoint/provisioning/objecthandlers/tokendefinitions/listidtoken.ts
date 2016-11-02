import {TokenDefinition} from "./tokendefinition";
import {TokensUtils} from "./tokensutils";

export class ListIdToken extends TokenDefinition {
    private _listId: string = null;

    constructor(web: SP.Web, name: string, listId: SP.Guid) {
        super(web, ["{{listid:" + TokensUtils.escapeRegExp(name) + "}}"]);
        this._listId = listId.toString();
    }

    public getReplaceValue(): Promise<string> {
        return new Promise<string>((resolve, reject) => {
            if (this.cacheValue != null && this.cacheValue !== "") {
                this.cacheValue = this._listId;
            }
            resolve(this.cacheValue);
        });
    }
}
