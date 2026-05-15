import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
export declare class Visual implements IVisual {
    private host;
    private root;
    private toolbar;
    private scrollWrap;
    private exportHeaders;
    private exportRows;
    constructor(options: VisualConstructorOptions);
    update(options: VisualUpdateOptions): void;
    private render;
    private exportCSV;
}
