import {TokenDefinition} from "./tokendefinition";
import {TokensUtils} from "./tokensutils";

export class ListUrlToken extends TokenDefinition {
    private _listUrl: string = null;

    constructor(web: SP.Web, name: string, listUrl: string) {
        super(web, ["<<listurl:" + TokensUtils.escapeRegExp(name) + ">>"]);
        this._listUrl = listUrl;
    }

    public getReplaceValue(): Promise<string> {
        return new Promise<string>((resolve, reject) => {
            if (this.cacheValue == null || this.cacheValue === "") {
                this.cacheValue = this._listUrl;
            }
            resolve(this.cacheValue);
        });
    }
}
