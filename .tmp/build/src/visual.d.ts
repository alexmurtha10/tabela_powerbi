import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
export declare class Visual implements IVisual {
    private host;
    private root;
    private toolbar;
    private scrollWrap;
    private overlay;
    private modal;
    private modalTitle;
    private modalBody;
    private exportHeaders;
    private exportRows;
    constructor(options: VisualConstructorOptions);
    private openModal;
    private closeModal;
    update(options: VisualUpdateOptions): void;
    private render;
    private exportCSV;
    private exportLocal;
}
