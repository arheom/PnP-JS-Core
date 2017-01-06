import {TokenDefinition} from "./tokendefinition";
import {TokensUtils} from "./tokensutils";

export class ContentTypeIdToken extends TokenDefinition {
    private _contentTypeId: string = null;

    constructor(web: SP.Web, name: string, contenttypeid: string) {
        super(web, ["<<contenttypeid:" + TokensUtils.escapeRegExp(name) + ">>"]);
        this._contentTypeId = contenttypeid;
    }

    public getReplaceValue(): Promise<string> {
        return new Promise<string>((resolve, reject) => {
            if (this.cacheValue == null || this.cacheValue === "") {
                this.cacheValue = this._contentTypeId;
            }
            resolve(this.cacheValue);
        });
    }
}
