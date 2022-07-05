<?php

namespace omnisoftory\PhpOffice\PhpSpreadsheet\Writer;

use \PhpOffice\PhpSpreadsheet\Calculation\Calculation;
use \PhpOffice\PhpSpreadsheet\Calculation\Functions;
use \PhpOffice\PhpSpreadsheet\HashTable;
use \PhpOffice\PhpSpreadsheet\Shared\File;
use \PhpOffice\PhpSpreadsheet\Spreadsheet;
// use omnisoftory\PhpOffice\PhpSpreadsheet\Spreadsheet;
use \PhpOffice\PhpSpreadsheet\Worksheet\Drawing as WorksheetDrawing;
use \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing;
use \PhpOffice\PhpSpreadsheet\Writer\Exception as WriterException;
//use \PhpOffice\PhpSpreadsheet\Writer\Xlsx\Chart;
use \PhpOffice\PhpSpreadsheet\Writer\Xlsx\Comments;
use \PhpOffice\PhpSpreadsheet\Writer\Xlsx\ContentTypes;
use \PhpOffice\PhpSpreadsheet\Writer\Xlsx\DocProps;
use \PhpOffice\PhpSpreadsheet\Writer\Xlsx\Drawing;
// use \PhpOffice\PhpSpreadsheet\Writer\Xlsx\Rels;
use omnisoftory\PhpOffice\PhpSpreadsheet\Writer\Xlsx\Rels;
use \PhpOffice\PhpSpreadsheet\Writer\Xlsx\RelsRibbon;
use \PhpOffice\PhpSpreadsheet\Writer\Xlsx\RelsVBA;
use \PhpOffice\PhpSpreadsheet\Writer\Xlsx\StringTable;
use \PhpOffice\PhpSpreadsheet\Writer\Xlsx\Style;
use \PhpOffice\PhpSpreadsheet\Writer\Xlsx\Theme;
use \PhpOffice\PhpSpreadsheet\Writer\Xlsx\Table;
// use \PhpOffice\PhpSpreadsheet\Writer\Xlsx\Workbook;
use omnisoftory\PhpOffice\PhpSpreadsheet\Writer\Xlsx\Workbook;
use \PhpOffice\PhpSpreadsheet\Writer\Xlsx\Worksheet;
use \ZipArchive;
use omnisoftory\PhpOffice\PhpSpreadsheet\Xlsx\PivotCacheDefinition;
use omnisoftory\PhpOffice\PhpSpreadsheet\Xlsx\PivotTable;
use omnisoftory\PhpOffice\PhpSpreadsheet\Xlsx\PivotCacheRecords;
use omnisoftory\PhpOffice\PhpSpreadsheet\Charts\ChartWriter;

require_once __DIR__ . "/Xlsx/Rels.php";
require_once __DIR__ . "/Spreadsheet.php";
require_once __DIR__ . "/Xlsx/Workbook.php";
require_once __DIR__ . "/Charts/ChartWriter.php";

class ZipArchiveX extends ZipArchive
{
	public function formattedXml( $contents )
	{
		$xml = new \DOMDocument();
		$xml->preserveWhiteSpace = false;
		$xml->formatOutput = true;
		$xml->loadXML($contents);
		return trim( $xml->saveXML() );
	}
	public function addFormattedXml( $localname, $contents )
	{
		// return $this->addFromString( $localname, $contents );
		return $this->addFromString( $localname, $this->formattedXml( $contents ) );
	}
}

class Xlsx extends \PhpOffice\PhpSpreadsheet\Writer\Xlsx
{
    /**
     * Office2003 compatibility.
     *
     * @var bool
     */
    private $office2003compatibility = false;


    /**
     * Private Spreadsheet.
     *
     * @var Spreadsheet
     */
    private $spreadSheet;

    /**
     * Private string table.
     *
     * @var string[]
     */
    private $stringTable = [];

    /**
     * Private unique Conditional HashTable.
     *
     * @var HashTable
     */
    private $stylesConditionalHashTable;

    /**
     * Private unique Style HashTable.
     *
     * @var HashTable
     */
    private $styleHashTable;

    /**
     * Private unique Fill HashTable.
     *
     * @var HashTable
     */
    private $fillHashTable;

    /**
     * Private unique \PhpOffice\PhpSpreadsheet\Style\Font HashTable.
     *
     * @var HashTable
     */
    private $fontHashTable;

    /**
     * Private unique Borders HashTable.
     *
     * @var HashTable
     */
    private $bordersHashTable;

    /**
     * Private unique NumberFormat HashTable.
     *
     * @var HashTable
     */
    private $numFmtHashTable;

    /**
     * Private unique \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet\BaseDrawing HashTable.
     *
     * @var HashTable
     */
    private $drawingHashTable;
    /**
     * @var Chart
     */
    private $writerPartChart;
    
    /**
     * @var Comments
     */
    private $writerPartComments;
    
    /**
     * @var ContentTypes
     */
    private $writerPartContentTypes;
    
    /**
     * @var DocProps
     */
    private $writerPartDocProps;
    
    /**
     * @var Drawing
     */
    private $writerPartDrawing;
    
    /**
     * @var Rels
     */
    private $writerPartRels;
    
    /**
     * @var RelsRibbon
     */
    private $writerPartRelsRibbon;
    
    /**
     * @var RelsVBA
     */
    private $writerPartRelsVBA;
    
    /**
     * @var StringTable
     */
    private $writerPartStringTable;
    
    /**
     * @var Style
     */
    private $writerPartStyle;
    
    /**
     * @var Theme
     */
    private $writerPartTheme;
    
    /**
     * @var Table
     */
    private $writerPartTable;
    
    /**
     * @var Workbook
     */
    private $writerPartWorkbook;
    
    /**
     * @var Worksheet
     */
    private $writerPartWorksheet;
    
    /**
     * Create a new Xlsx Writer.
     *
     * @param Spreadsheet $spreadsheet
     * PHP 7.x does not recognize a constructor using a descendant class as valid
     */
    public function __construct(\PhpOffice\PhpSpreadsheet\Spreadsheet $spreadsheet)
    {
        // Assign PhpSpreadsheet
        $this->setSpreadsheet($spreadsheet);
        
        $this->writerPartStringTable = new StringTable($this);
        $this->writerPartContentTypes = new ContentTypes($this);
        $this->writerPartDocProps = new DocProps($this);
        $this->writerPartTheme = new Theme($this);
        $this->writerPartRels = new Rels($this);
        $this->writerPartStyle = new Style($this);
        $this->writerPartWorkbook = new Workbook($this);
        $this->writerPartWorksheet = new Worksheet($this);
        $this->writerPartDrawing = new Drawing($this);
        $this->writerPartChart = new ChartWriter($this);
        $this->writerPartComments = new Comments($this);
        $this->writerPartRelsRibbon = new RelsRibbon($this);
        $this->writerPartRelsVBA = new RelsVBA($this);
        
        $this->writerPartTable = new Table($this);

        $hashTablesArray = ['stylesConditionalHashTable', 'fillHashTable', 'fontHashTable',
            'bordersHashTable', 'numFmtHashTable', 'drawingHashTable',
            'styleHashTable',
        ];

        // Set HashTable variables
        foreach ($hashTablesArray as $tableName) {
            $this->$tableName = new HashTable();
        }
    }

    
    public function getWriterPartChart(): \PhpOffice\PhpSpreadsheet\Writer\Xlsx\Chart
    {
        return $this->writerPartChart;
    }
    
    public function getWriterPartComments(): \PhpOffice\PhpSpreadsheet\Writer\Xlsx\Comments
    {
        return $this->writerPartComments;
    }
    
    public function getWriterPartContentTypes(): \PhpOffice\PhpSpreadsheet\Writer\Xlsx\ContentTypes
    {
        return $this->writerPartContentTypes;
    }
    
    public function getWriterPartDocProps(): \PhpOffice\PhpSpreadsheet\Writer\Xlsx\DocProps
    {
        return $this->writerPartDocProps;
    }
    
    public function getWriterPartDrawing(): \PhpOffice\PhpSpreadsheet\Writer\Xlsx\Drawing
    {
        return $this->writerPartDrawing;
    }
    
    public function getWriterPartRels(): \PhpOffice\PhpSpreadsheet\Writer\Xlsx\Rels
    {
        return $this->writerPartRels;
    }
    
    public function getWriterPartRelsRibbon(): \PhpOffice\PhpSpreadsheet\Writer\Xlsx\RelsRibbon
    {
        return $this->writerPartRelsRibbon;
    }
    
    public function getWriterPartRelsVBA(): \PhpOffice\PhpSpreadsheet\Writer\Xlsx\RelsVBA
    {
        return $this->writerPartRelsVBA;
    }
    
    public function getWriterPartStringTable(): \PhpOffice\PhpSpreadsheet\Writer\Xlsx\StringTable
    {
        return $this->writerPartStringTable;
    }
    
    public function getWriterPartStyle(): \PhpOffice\PhpSpreadsheet\Writer\Xlsx\Style
    {
        return $this->writerPartStyle;
    }
    
    public function getWriterPartTheme(): \PhpOffice\PhpSpreadsheet\Writer\Xlsx\Theme
    {
        return $this->writerPartTheme;
    }
    
    public function getWriterPartTable(): \PhpOffice\PhpSpreadsheet\Writer\Xlsx\Table
    {
        return $this->writerPartTable;
    }
    
    public function getWriterPartWorkbook(): \PhpOffice\PhpSpreadsheet\Writer\Xlsx\Workbook
    {
        return $this->writerPartWorkbook;
    }
    
    public function getWriterPartWorksheet(): \PhpOffice\PhpSpreadsheet\Writer\Xlsx\Worksheet
    {
        return $this->writerPartWorksheet;
    }

    /**
     * Save PhpSpreadsheet to file.
     *
     * @param string $pFilename
     *
     * @throws WriterException
     */
    public function save($pFilename, int $flags = 0) : void
    {
        if ($this->spreadSheet !== null) {
            // garbage collect
            $this->spreadSheet->garbageCollect();

            // If $pFilename is php://output or php://stdout, make it a temporary file...
            $originalFilename = $pFilename;
            if (strtolower($pFilename) == 'php://output' || strtolower($pFilename) == 'php://stdout') {
                $pFilename = @tempnam(File::sysGetTempDir(), 'phpxltmp');
                if ($pFilename == '') {
                    $pFilename = $originalFilename;
                }
            }

            $saveDebugLog = Calculation::getInstance($this->spreadSheet)->getDebugLog()->getWriteDebugLog();
            Calculation::getInstance($this->spreadSheet)->getDebugLog()->setWriteDebugLog(false);
            $saveDateReturnType = Functions::getReturnDateType();
            Functions::setReturnDateType(Functions::RETURNDATE_EXCEL);

            // Create string lookup table
            $this->stringTable = [];
            for ($i = 0; $i < $this->spreadSheet->getSheetCount(); ++$i) {
                $this->stringTable = $this->getWriterPartStringTable()->createStringTable($this->spreadSheet->getSheet($i), $this->stringTable);
            }

            // Create styles dictionaries
            $this->styleHashTable->addFromSource($this->getWriterPartStyle()->allStyles($this->spreadSheet));
            $this->stylesConditionalHashTable->addFromSource($this->getWriterPartStyle()->allConditionalStyles($this->spreadSheet));
            $this->fillHashTable->addFromSource($this->getWriterPartStyle()->allFills($this->spreadSheet));
            $this->fontHashTable->addFromSource($this->getWriterPartStyle()->allFonts($this->spreadSheet));
            $this->bordersHashTable->addFromSource($this->getWriterPartStyle()->allBorders($this->spreadSheet));
            $this->numFmtHashTable->addFromSource($this->getWriterPartStyle()->allNumberFormats($this->spreadSheet));

            // Create drawing dictionary
            $this->drawingHashTable->addFromSource($this->getWriterPartDrawing()->allDrawings($this->spreadSheet));

            $zip = new ZipArchiveX();

            if (file_exists($pFilename)) {
                unlink($pFilename);
            }
            // Try opening the ZIP file
            if ($zip->open($pFilename, ZipArchive::OVERWRITE) !== true) {
                if ($zip->open($pFilename, ZipArchive::CREATE) !== true) {
                    throw new WriterException('Could not open ' . $pFilename . ' for writing.');
                }
            }

            // Add [Content_Types].xml to ZIP file
			$unparsedLoadedData = $this->spreadSheet->getUnparsedLoadedData();
			$definitions = $this->spreadSheet->pivotCacheDefinitionCollection;
			if ( $definitions && $definitions->hasPivotCacheDefinitions() )
			{
				foreach ( $definitions as $path => /** @var PivotCacheDefinition $definition */ $definition )
				{
					$unparsedLoadedData['override_content_types'][ "/$path" ] = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml";
					$zip->addFormattedXml( $definition->path, $definition->xml );
					$zip->addFormattedXml('xl/pivotCache/_rels/' . basename( $definition->path ) . '.rels', $this->getWriterPartRels()->writeCacheRelationships($this->spreadSheet, $definition));
				}
			}
			$pivotTables = $this->spreadSheet->pivotTables;
			if ( $pivotTables && $pivotTables->hasPivotTables() )
			{
				foreach ( $pivotTables as $path => /** @var PivotTable $table */ $table )
				{
					$unparsedLoadedData['override_content_types'][ "/$path" ] = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml";
					$zip->addFormattedXml( $table->path, $table->xml );
					$zip->addFormattedXml('xl/pivotTables/_rels/' . basename( $table->path ) . '.rels', $this->getWriterPartRels()->writePivotTableRelationships($this->spreadSheet, $table));
				}
			}
			$records = $this->spreadSheet->pivotCacheRecordsCollection;
			if ( $records && $records->hasPivotCacheRecords() )
			{
				foreach ( $records as $path => /** @var PivotCacheRecords $records */ $records )
				{
					$unparsedLoadedData['override_content_types'][ "/$path" ] = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml";
					$zip->addFormattedXml( $records->path, $records->xml );
				}
			}
			$this->spreadSheet->setUnparsedLoadedData( $unparsedLoadedData );

            $zip->addFormattedXml('[Content_Types].xml', $this->getWriterPartContentTypes()->writeContentTypes($this->spreadSheet, $this->includeCharts));

            //if hasMacros, add the vbaProject.bin file, Certificate file(if exists)
            if ($this->spreadSheet->hasMacros()) {
                $macrosCode = $this->spreadSheet->getMacrosCode();
                if ($macrosCode !== null) {
                    // we have the code ?
                    $zip->addFormattedXml('xl/vbaProject.bin', $macrosCode); //allways in 'xl', allways named vbaProject.bin
                    if ($this->spreadSheet->hasMacrosCertificate()) {
                        //signed macros ?
                        // Yes : add the certificate file and the related rels file
                        $zip->addFormattedXml('xl/vbaProjectSignature.bin', $this->spreadSheet->getMacrosCertificate());
                        $zip->addFormattedXml('xl/_rels/vbaProject.bin.rels', $this->getWriterPartRelsVBA()->writeVBARelationships($this->spreadSheet));
                    }
                }
            }
            //a custom UI in this workbook ? add it ("base" xml and additional objects (pictures) and rels)
            if ($this->spreadSheet->hasRibbon()) {
                $tmpRibbonTarget = $this->spreadSheet->getRibbonXMLData('target');
                $zip->addFormattedXml($tmpRibbonTarget, $this->spreadSheet->getRibbonXMLData('data'));
                if ($this->spreadSheet->hasRibbonBinObjects()) {
                    $tmpRootPath = dirname($tmpRibbonTarget) . '/';
                    $ribbonBinObjects = $this->spreadSheet->getRibbonBinObjects('data'); //the files to write
                    foreach ($ribbonBinObjects as $aPath => $aContent) {
                        $zip->addFormattedXml($tmpRootPath . $aPath, $aContent);
                    }
                    //the rels for files
                    $zip->addFormattedXml($tmpRootPath . '_rels/' . basename($tmpRibbonTarget) . '.rels', $this->getWriterPartRelsRibbon()->writeRibbonRelationships($this->spreadSheet));
                }
            }

            // Add relationships to ZIP file
            $zip->addFormattedXml('_rels/.rels', $this->getWriterPartRels()->writeRelationships($this->spreadSheet));
            // BMS 2018-10-20 Added to writeWorkbookRelationships
            $zip->addFormattedXml('xl/_rels/workbook.xml.rels', $this->getWriterPartRels()->writeWorkbookRelationships($this->spreadSheet));

            // Add document properties to ZIP file
            $zip->addFormattedXml('docProps/app.xml', $this->getWriterPartDocProps()->writeDocPropsApp($this->spreadSheet));
            $zip->addFormattedXml('docProps/core.xml', $this->getWriterPartDocProps()->writeDocPropsCore($this->spreadSheet));
            $customPropertiesPart = $this->getWriterPartDocProps()->writeDocPropsCustom($this->spreadSheet);
            if ($customPropertiesPart !== null) {
                $zip->addFormattedXml('docProps/custom.xml', $customPropertiesPart);
            }

            // Add theme to ZIP file
            $zip->addFormattedXml('xl/theme/theme1.xml', $this->getWriterPartTheme()->writeTheme($this->spreadSheet));

            // Add string table to ZIP file
            $zip->addFormattedXml('xl/sharedStrings.xml', $this->getWriterPartStringTable()->writeStringTable($this->stringTable));

            // Add styles to ZIP file
            $zip->addFormattedXml('xl/styles.xml', $this->getWriterPartStyle()->writeStyles($this->spreadSheet));

            $chartCount = 0;
            // Add worksheets
            for ($i = 0; $i < $this->spreadSheet->getSheetCount(); ++$i) {
                $zip->addFormattedXml('xl/worksheets/sheet' . ($i + 1) . '.xml', $this->getWriterPartWorksheet()->writeWorksheet($this->spreadSheet->getSheet($i), $this->stringTable, $this->includeCharts));
                if ($this->includeCharts) {
                    $charts = $this->spreadSheet->getSheet($i)->getChartCollection();
                    if (count($charts) > 0) {
                        foreach ($charts as $chart) {
                            $zip->addFormattedXml('xl/charts/chart' . ($chartCount + 1) . '.xml', $this->getWriterPartChart()->writeChart($chart, $this->preCalculateFormulas));
                            ohTrace(OH_TRACE_RELEASE, "LEA chartxml: ". $this->getWriterPartChart()->writeChart($chart, $this->preCalculateFormulas),__FILE__,__LINE__,0);
                            ++$chartCount;
                        }
                    }
                }
            }

            $chartRef1 = 0;
            // Add worksheet relationships (drawings, ...)
            for ($i = 0; $i < $this->spreadSheet->getSheetCount(); ++$i) {
                // Add relationships
                $zip->addFormattedXml('xl/worksheets/_rels/sheet' . ($i + 1) . '.xml.rels', $this->getWriterPartRels()->writeWorksheetRelationships($this->spreadSheet->getSheet($i), ($i + 1), $this->includeCharts));

                // Add unparsedLoadedData
                $sheetCodeName = $this->spreadSheet->getSheet($i)->getCodeName();
                $unparsedLoadedData = $this->spreadSheet->getUnparsedLoadedData();
                if (isset($unparsedLoadedData['sheets'][$sheetCodeName]['ctrlProps'])) {
                    foreach ($unparsedLoadedData['sheets'][$sheetCodeName]['ctrlProps'] as $ctrlProp) {
                        $zip->addFormattedXml($ctrlProp['filePath'], $ctrlProp['content']);
                    }
                }
                if (isset($unparsedLoadedData['sheets'][$sheetCodeName]['printerSettings'])) {
                    foreach ($unparsedLoadedData['sheets'][$sheetCodeName]['printerSettings'] as $ctrlProp) {
                        $zip->addFormattedXml($ctrlProp['filePath'], $ctrlProp['content']);
                    }
                }

                $drawings = $this->spreadSheet->getSheet($i)->getDrawingCollection();
                $drawingCount = count($drawings);
                if ($this->includeCharts) {
                    $chartCount = $this->spreadSheet->getSheet($i)->getChartCount();
                }

                // Add drawing and image relationship parts
                if (($drawingCount > 0) || ($chartCount > 0)) {
                    // Drawing relationships
                    $zip->addFormattedXml('xl/drawings/_rels/drawing' . ($i + 1) . '.xml.rels', $this->getWriterPartRels()->writeDrawingRelationships($this->spreadSheet->getSheet($i), $chartRef1, $this->includeCharts));

                    // Drawings
                    $zip->addFormattedXml('xl/drawings/drawing' . ($i + 1) . '.xml', $this->getWriterPartDrawing()->writeDrawings($this->spreadSheet->getSheet($i), $this->includeCharts));
                } elseif (isset($unparsedLoadedData['sheets'][$sheetCodeName]['drawingAlternateContents'])) {
                    // Drawings
                    $zip->addFormattedXml('xl/drawings/drawing' . ($i + 1) . '.xml', $this->getWriterPartDrawing()->writeDrawings($this->spreadSheet->getSheet($i), $this->includeCharts));
                }

                // Add comment relationship parts
                if (count($this->spreadSheet->getSheet($i)->getComments()) > 0) {
                    // VML Comments
                    $zip->addFormattedXml('xl/drawings/vmlDrawing' . ($i + 1) . '.vml', $this->getWriterPartComments()->writeVMLComments($this->spreadSheet->getSheet($i)));

                    // Comments
                    $zip->addFormattedXml('xl/comments' . ($i + 1) . '.xml', $this->getWriterPartComments()->writeComments($this->spreadSheet->getSheet($i)));
                }

                // Add unparsed relationship parts
                if (isset($unparsedLoadedData['sheets'][$this->spreadSheet->getSheet($i)->getCodeName()]['vmlDrawings'])) {
                    foreach ($unparsedLoadedData['sheets'][$this->spreadSheet->getSheet($i)->getCodeName()]['vmlDrawings'] as $vmlDrawing) {
                        $zip->addFormattedXml($vmlDrawing['filePath'], $vmlDrawing['content']);
                    }
                }

                // Add header/footer relationship parts
                if (count($this->spreadSheet->getSheet($i)->getHeaderFooter()->getImages()) > 0) {
                    // VML Drawings
                    $zip->addFormattedXml('xl/drawings/vmlDrawingHF' . ($i + 1) . '.vml', $this->getWriterPartDrawing()->writeVMLHeaderFooterImages($this->spreadSheet->getSheet($i)));

                    // VML Drawing relationships
                    $zip->addFormattedXml('xl/drawings/_rels/vmlDrawingHF' . ($i + 1) . '.vml.rels', $this->getWriterPartRels()->writeHeaderFooterDrawingRelationships($this->spreadSheet->getSheet($i)));

                    // Media
                    foreach ($this->spreadSheet->getSheet($i)->getHeaderFooter()->getImages() as $image) {
                        $zip->addFormattedXml('xl/media/' . $image->getIndexedFilename(), file_get_contents($image->getPath()));
                    }
                }
            }

            // Add media
            for ($i = 0; $i < $this->getDrawingHashTable()->count(); ++$i) {
                if ($this->getDrawingHashTable()->getByIndex($i) instanceof WorksheetDrawing) {
                    $imageContents = null;
                    $imagePath = $this->getDrawingHashTable()->getByIndex($i)->getPath();
                    if (strpos($imagePath, 'zip://') !== false) {
                        $imagePath = substr($imagePath, 6);
                        $imagePathSplitted = explode('#', $imagePath);

                        $imageZip = new ZipArchive();
                        $imageZip->open($imagePathSplitted[0]);
                        $imageContents = $imageZip->getFromName($imagePathSplitted[1]);
                        $imageZip->close();
                        unset($imageZip);
                    } else {
                        $imageContents = file_get_contents($imagePath);
                    }

                    $zip->addFormattedXml('xl/media/' . str_replace(' ', '_', $this->getDrawingHashTable()->getByIndex($i)->getIndexedFilename()), $imageContents);
                } elseif ($this->getDrawingHashTable()->getByIndex($i) instanceof MemoryDrawing) {
                    ob_start();
                    call_user_func(
                        $this->getDrawingHashTable()->getByIndex($i)->getRenderingFunction(),
                        $this->getDrawingHashTable()->getByIndex($i)->getImageResource()
                    );
                    $imageContents = ob_get_contents();
                    ob_end_clean();

                    $zip->addFormattedXml('xl/media/' . str_replace(' ', '_', $this->getDrawingHashTable()->getByIndex($i)->getIndexedFilename()), $imageContents);
                }
            }

            // Add workbook to ZIP file
            // BMS 2018-10-20
            $zip->addFormattedXml('xl/workbook.xml', $this->getWriterPartWorkbook()->writeWorkbook($this->spreadSheet, $this->preCalculateFormulas));

            Functions::setReturnDateType($saveDateReturnType);
            Calculation::getInstance($this->spreadSheet)->getDebugLog()->setWriteDebugLog($saveDebugLog);

            // Close file
            if ($zip->close() === false) {
                throw new WriterException("Could not close zip file $pFilename.");
            }

            // If a temporary file was used, copy it to the correct file stream
            if ($originalFilename != $pFilename) {
                if (copy($pFilename, $originalFilename) === false) {
                    throw new WriterException("Could not copy temporary zip file $pFilename to $originalFilename.");
                }
                @unlink($pFilename);
            }
        } else {
            throw new WriterException('PhpSpreadsheet object unassigned.');
        }
    }

    /**
     * Get Spreadsheet object.
     *
     * @throws WriterException
     *
     * @return Spreadsheet
     */
    public function getSpreadsheet()
    {
        if ($this->spreadSheet !== null) {
            return $this->spreadSheet;
        }

        throw new WriterException('No Spreadsheet object assigned.');
    }

    /**
     * Set Spreadsheet object.
     *
     * @param Spreadsheet $spreadsheet PhpSpreadsheet object
     *
     * @return Xlsx
     */
    public function setSpreadsheet(\PhpOffice\PhpSpreadsheet\Spreadsheet $spreadsheet)
    {
        $this->spreadSheet = $spreadsheet;

        return $this;
    }

    /**
     * Get string table.
     *
     * @return string[]
     */
    public function getStringTable()
    {
        return $this->stringTable;
    }

    /**
     * Get Style HashTable.
     *
     * @return HashTable
     */
    public function getStyleHashTable()
    {
        return $this->styleHashTable;
    }

    /**
     * Get Conditional HashTable.
     *
     * @return HashTable
     */
    public function getStylesConditionalHashTable()
    {
        return $this->stylesConditionalHashTable;
    }

    /**
     * Get Fill HashTable.
     *
     * @return HashTable
     */
    public function getFillHashTable()
    {
        return $this->fillHashTable;
    }

    /**
     * Get \PhpOffice\PhpSpreadsheet\Style\Font HashTable.
     *
     * @return HashTable
     */
    public function getFontHashTable()
    {
        return $this->fontHashTable;
    }

    /**
     * Get Borders HashTable.
     *
     * @return HashTable
     */
    public function getBordersHashTable()
    {
        return $this->bordersHashTable;
    }

    /**
     * Get NumberFormat HashTable.
     *
     * @return HashTable
     */
    public function getNumFmtHashTable()
    {
        return $this->numFmtHashTable;
    }

    /**
     * Get \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet\BaseDrawing HashTable.
     *
     * @return HashTable
     */
    public function getDrawingHashTable()
    {
        return $this->drawingHashTable;
    }

    /**
     * Get Office2003 compatibility.
     *
     * @return bool
     */
    public function getOffice2003Compatibility()
    {
        return $this->office2003compatibility;
    }

    /**
     * Set Office2003 compatibility.
     *
     * @param bool $pValue Office2003 compatibility?
     *
     * @return Xlsx
     */
    public function setOffice2003Compatibility($pValue)
    {
        $this->office2003compatibility = $pValue;

        return $this;
    }
}
