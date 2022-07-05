<?php

namespace lyquidity\xbrl_validate\PhpOffice\PhpSpreadsheet\Charts;

use PhpOffice\PhpSpreadsheet\Chart\Axis;
use PhpOffice\PhpSpreadsheet\Chart\DataSeries;
use PhpOffice\PhpSpreadsheet\Chart\GridLines;
use PhpOffice\PhpSpreadsheet\Chart\Legend;
use PhpOffice\PhpSpreadsheet\Chart\PlotArea;
use PhpOffice\PhpSpreadsheet\Chart\Title;


class Chart extends \PhpOffice\PhpSpreadsheet\Chart\Chart
{
    
    private $_hasPivotSource = false;
    private $_pivotSourceName = "";
    private $_fmtId;
    private $_extLstXMLElement;
    private $_view3DXMLElement;
/**
     * Create a new Chart.
     *
     * @param mixed $name
     * @param mixed $plotVisibleOnly
     * @param string $displayBlanksAs
     */
    public function __construct($name, ?Title $title = null, ?Legend $legend = null, ?PlotArea $plotArea = null, $plotVisibleOnly = true, $displayBlanksAs = DataSeries::EMPTY_AS_GAP, ?Title $xAxisLabel = null, ?Title $yAxisLabel = null, ?Axis $xAxis = null, ?Axis $yAxis = null, ?GridLines $majorGridlines = null, ?GridLines $minorGridlines = null)
    {
        parent::__construct($name,$title,$legend,$plotArea,$plotVisibleOnly,$displayBlanksAs,$xAxisLabel,$yAxisLabel,$xAxis,$yAxis,$majorGridlines,$minorGridlines);
        
        
    }
    /**
     * @return boolean
     */
    public function getHasPivotSource()
    {
        return $this->_hasPivotSource;
    }
    
    /**
     * @param boolean $_hasPivotSource
     */
    public function setHasPivotSource($_hasPivotSource)
    {
        $this->_hasPivotSource = $_hasPivotSource;
    }
    
    /**
     * @return string
     */
    public function getPivotSourceName()
    {
        return $this->_pivotSourceName;
    }
    
    /**
     * @param string $_pivotSourceName
     */
    public function setPivotSourceName($_pivotSourceName)
    {
        $this->_pivotSourceName = $_pivotSourceName;
    }
    
    /**
     * @return mixed
     */
    public function getFmtId()
    {
        return $this->_fmtId;
    }
    
    /**
     * @param mixed $_fmtId
     */
    public function setFmtId($_fmtId)
    {
        $this->_fmtId = $_fmtId;
    }
    /**
     * @return mixed
     */
    public function getExtLstXMLElement()
    {
        return $this->_extLstXMLElement;
    }

    /**
     * @param mixed $_extLstXMLElement
     */
    public function setExtLstXMLElement($_extLstXMLElement)
    {
        $this->_extLstXMLElement = $_extLstXMLElement;
    }
    /**
     * @return mixed
     */
    public function getView3DXMLElement()
    {
        return $this->_view3DXMLElement;
    }

    /**
     * @param mixed $_view3DXMLElement
     */
    public function setView3DXMLElement($_view3DXMLElement)
    {
        $this->_view3DXMLElement = $_view3DXMLElement;
    }


    
    
}
