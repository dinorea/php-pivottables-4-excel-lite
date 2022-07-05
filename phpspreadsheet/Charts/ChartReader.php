<?php

namespace omnisoftory\PhpOffice\PhpSpreadsheet\Charts;
require_once __DIR__ . "/Chart.php";
use SimpleXMLElement;


class ChartReader extends \PhpOffice\PhpSpreadsheet\Reader\Xlsx\Chart
{

    /**
     * @param string $chartName
     *
     * @return
     */
    public static function readChart(SimpleXMLElement $chartElements, $chartName)
    {
        $parentChart = parent::readChart($chartElements, $chartName);
                
        $namespacesChartMeta = $chartElements->getNamespaces(true);
        $chartElementsC = $chartElements->children($namespacesChartMeta['c']);

        $pivotChart =  new \omnisoftory\PhpOffice\PhpSpreadsheet\Charts\Chart($chartName, $parentChart->getTitle(), $parentChart->getLegend(), $parentChart->getPlotArea(), $parentChart->getPlotVisibleOnly(), $parentChart->getDisplayBlanksAs(), $parentChart->getXAxisLabel(), $parentChart->getYAxisLabel());
        foreach ($chartElementsC as $chartElementKey => $chartElement) {
            switch ($chartElementKey) {
                case 'pivotSource':
                    $pivotChart->setHasPivotSource(true);
                    foreach ($chartElement as $chartDetailsKey => $chartDetails) {
                        switch ($chartDetailsKey) {
                            case 'fmtId':
                                $pivotChart->setFmtId(self::getAttribute($chartDetails, 'val', 'int'));
                                break;
                            case 'name':
                                $pivotChart->setPivotSourceName($chartDetails->__toString());
                                break;
                        }
                    }
                    break;
                case 'extLst':
                    $pivotChart->setExtLstXMLElement($chartElement);
                    break;
                case 'chart' :
                    foreach ( $chartElement as $chartDetailsKey => $chartDetails )
                    {
                        switch ($chartDetailsKey)
                        {
                            case 'view3D' :
                                $pivotChart->setView3DXMLElement($chartDetails);
                                break;
                        }
                    }
                    break;                       
            }
        }
        return $pivotChart;
    }
    
    
    /**
     * @param string $name
     * @param string $format
     *
     * @return null|bool|float|int|string
     */
    private static function getAttribute(SimpleXMLElement $component, $name, $format)
    {
        $attributes = $component->attributes();
        if (isset($attributes[$name])) {
            if ($format == 'string') {
                return (string) $attributes[$name];
            } elseif ($format == 'integer') {
                return (int) $attributes[$name];
            } elseif ($format == 'boolean') {
                $value = (string) $attributes[$name];
                
                return $value === 'true' || $value === '1';
            }
            
            return (float) $attributes[$name];
        }
        
        return null;
    }


}
