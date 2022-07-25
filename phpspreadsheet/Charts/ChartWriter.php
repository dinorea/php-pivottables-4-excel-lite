<?php

namespace omnisoftory\PhpOffice\PhpSpreadsheet\Charts;

use PhpOffice\PhpSpreadsheet\Chart\Layout;
use PhpOffice\PhpSpreadsheet\Chart\Title;
use PhpOffice\PhpSpreadsheet\Shared\XMLWriter;
require_once __DIR__ . "/Chart.php";


class ChartWriter extends \PhpOffice\PhpSpreadsheet\Writer\Xlsx\Chart
{
    
    
    /**
     * Write charts to XML format.
     *
     * @param mixed $calculateCellValues
     *
     * @return string XML Output
     */
    public function writeChart(\PhpOffice\PhpSpreadsheet\Charts\Chart $chart, $calculateCellValues = true)
    { /**$filename, **/
        $filename = "test";
        $chartXml = parent::writeChart($chart,$calculateCellValues);
        if (!$chart->getHasPivotSource()) 
        {
            return $chartXml;
        }
        
        if (strpos($chartXml, "<c:title>")) 
        {
            //replace title
            $startStr = substr($chartXml, 0, strpos($chartXml, "<c:title>"));
            $titleStr = $this->writeTitle($chart->getTitle());
            $endStr = substr($chartXml, strpos($chartXml, "</c:title>")+10);
            $chartXml = $startStr.$titleStr.$endStr;
            
        }
        
        if ($chart->getHasPivotSource()) 
        {
            //Add pivotSource to the xml before the chart element
            $startStr = substr($chartXml, 0, strpos($chartXml, "<c:chart>"));
            $pivotStr = $this->writePivotSource($chart, $filename);
            $endStr = substr($chartXml, strpos($chartXml, "<c:chart>"));
            $chartXml = $startStr.$pivotStr.$endStr;
        }
        
        if ($chart->getExtLstXMLElement() != null) 
        {
            
            //Add extLst (pivotOptions) to the xml before the chart element
            $startStr = substr($chartXml, 0, strpos($chartXml, "</c:chartSpace>"));
            $extLstStr = $chart->getExtLstXMLElement()->asXML();
            $endStr = substr($chartXml, strpos($chartXml, "</c:chartSpace>"));
            $chartXml = $startStr.$extLstStr.$endStr;
            
        }
        
        if ( $chart->getView3DXMLElement() != null) {
            
            //Add view3D to the xml before the chart element
            $startStr = substr($chartXml, 0, strpos($chartXml, "</c:chart>"));
            $view3DStr = $chart->getView3DXMLElement()->asXML();
            $endStr = substr($chartXml, strpos($chartXml, "</c:chart>"));
            $chartXml = $startStr.$view3DStr.$endStr;
        }
        

        // Return
        return $chartXml;
    }

    /**
     * Write pivotSource.
     */
    private function writePivotSource($chart, $filename): string
    {
        // Create XML writer
        $objWriter = null;
        if ($this->getParentWriter()->getUseDiskCaching()) {
            $objWriter = new XMLWriter(XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
        } else {
            $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);
        }
        $objWriter->startElement('c:pivotSource');

        $sourceName = $chart->getPivotSourceName();
        $objWriter->startElement('c:name');
        $objWriter->writeRawData("[".$filename."]".substr($sourceName, strpos($sourceName, "]") + 1));
        $objWriter->endElement();

        $objWriter->startElement('c:fmtId');
        $objWriter->writeAttribute('val', $chart->getFmtId());
        $objWriter->endElement();

        $objWriter->endElement();
        
        return $objWriter->getData();
    }
    
    
    /**
     * Write Chart Title.
     */
    private function writeTitle(?Title $title = null): string
    {
        if ($title === null) {
            return "";
        }
        // Create XML writer
        $objWriter = null;
        if ($this->getParentWriter()->getUseDiskCaching()) {
            $objWriter = new XMLWriter(XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
        } else {
            $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);
        }
        
        $objWriter->startElement('c:title');
        
        if (!empty($title->getCaption()))
        {
            $objWriter->startElement('c:tx');
            $objWriter->startElement('c:rich');
            
            $objWriter->startElement('a:p');
            
            $caption = $title->getCaption();
            if ((is_array($caption)) && (count($caption) > 0)) {
                $caption = $caption[0];
            }
            $this->getParentWriter()->getWriterPartstringtable()->writeRichTextForCharts($objWriter, $caption, 'a');
            
            $objWriter->endElement();
            $objWriter->endElement();
            $objWriter->endElement();
        }
        
        $this->writeLayout($objWriter, $title->getLayout());
        
        $objWriter->startElement('c:overlay');
        $objWriter->writeAttribute('val', 0);
        $objWriter->endElement();
        
        $objWriter->endElement();
        return $objWriter->getData();
    }
    
    
    
    /**
     * Write Layout.
     */
    private function writeLayout(XMLWriter $objWriter, ?Layout $layout = null): void
    {
        $objWriter->startElement('c:layout');
        
        if ($layout !== null) {
            $objWriter->startElement('c:manualLayout');
            
            $layoutTarget = $layout->getLayoutTarget();
            if ($layoutTarget !== null) {
                $objWriter->startElement('c:layoutTarget');
                $objWriter->writeAttribute('val', $layoutTarget);
                $objWriter->endElement();
            }
            
            $xMode = $layout->getXMode();
            if ($xMode !== null) {
                $objWriter->startElement('c:xMode');
                $objWriter->writeAttribute('val', $xMode);
                $objWriter->endElement();
            }
            
            $yMode = $layout->getYMode();
            if ($yMode !== null) {
                $objWriter->startElement('c:yMode');
                $objWriter->writeAttribute('val', $yMode);
                $objWriter->endElement();
            }
            
            $x = $layout->getXPosition();
            if ($x !== null) {
                $objWriter->startElement('c:x');
                $objWriter->writeAttribute('val', $x);
                $objWriter->endElement();
            }
            
            $y = $layout->getYPosition();
            if ($y !== null) {
                $objWriter->startElement('c:y');
                $objWriter->writeAttribute('val', $y);
                $objWriter->endElement();
            }
            
            $w = $layout->getWidth();
            if ($w !== null) {
                $objWriter->startElement('c:w');
                $objWriter->writeAttribute('val', $w);
                $objWriter->endElement();
            }
            
            $h = $layout->getHeight();
            if ($h !== null) {
                $objWriter->startElement('c:h');
                $objWriter->writeAttribute('val', $h);
                $objWriter->endElement();
            }
            
            $objWriter->endElement();
        }
        
        $objWriter->endElement();
    }
    
}
