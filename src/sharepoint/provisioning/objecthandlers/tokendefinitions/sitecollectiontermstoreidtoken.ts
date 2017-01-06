import {TokenDefinition} from "./tokendefinition";

export class SiteCollectionTermStoreIdToken extends TokenDefinition {
    constructor(web: SP.Web) {
        super(web, ["~sitecollectiontermstoreid", "<<sitecollectiontermstoreid>>"]);
    }

    public getReplaceValue(): Promise<string> {
        return new Promise<string>((resolve, reject) => {
            if (this.cacheValue == null) {
                this.web.get_context().load(this.web, "Url");
                this.web.get_context().executeQueryAsync(() => {
                    let context = new SP.ClientContext(this.web.get_url());
                    let session = SP.Taxonomy.TaxonomySession.getTaxonomySession(context);
                    let termStores = session.get_termStores();
                    context.load(termStores);
                    context.executeQueryAsync(() => {
                        this.cacheValue = termStores.get_data()[0].get_id().toString();
                        resolve(this.cacheValue);
                    }, (sender, error) => { reject(error); });
                }, (sender, error) => { reject(error); });
            } else {
                resolve(this.cacheValue);
            }
        });
    }
}
