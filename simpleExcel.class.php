<?php
/**
 * simple excel export class
 */
namespace Util;

class simpleExcel {
	//数据转换临时目录
	const TMP_FILE_DIR = './';

	//最大读取字符串长度(2M)
	const MAX_LEN = 2097152;

	//导出文件名
	private $fileName;

	//创建临时存储文件
	private $tmpFile;

	//xls数据头映射
	private $headerMap;

	//xls数据头内容
	private $headerStr;

	//xls输出样式
	private $styleOption = array(
		//字体大小
		'fontSize'  => '12pt',
		//表头背景色
		'bgColor'   => '#72c5f5',
		//文本对齐方式
		'textAlign' => 'left'
	);

	//EXCEL导出格式映射
	private $styleMap = array(
		//文本
		'text' 		=>'@',
		//日期
		'date'		=>'yyyy-mm-dd', 
		//数字
		'number'	=>'#0',
		//浮点数
		'decimal'	=>'#0.00', 
		//百分比
		'percent'	=>'#0.00%',
	);

	public function __construct($styleOption = array()) {
		//设置xls样式配置项
		if (!empty($styleOption)) {
			$styleKeyMap = array_keys($this->styleOption);
			foreach ($styleOption as $key => $value) {
				if (!empty($value) && in_array($key, $styleKeyMap)) {
					$this->styleOption[$key] = $value;
				}
			}
		}
	}

	/**
	 * 设置导出文件名
	 * @param string $name 文件名
	 */
	public function setFileName($name) {
		if (!empty($name)) {
			$this->fileName = $name;
			$this->tmpFile = self::TMP_FILE_DIR . $name . '.tmp';
			if (file_exists($this->tmpFile)) {
				unlink($this->tmpFile);
			}
		}
	}

	/**
	 * 设置导出行数据头
	 *
	 * @param array $columnValue 行数据数组
	 */
	public function setXlsHeader($headerMap) {
		if (empty($headerValue)) {
			$this->headerMap = $headerMap;
			$this->headerStr = $this->formatXlsColumn($headerMap, true);
			//输出数据表头
			$fp = fopen($this->tmpFile, 'a+');
			fputs($fp, $this->headerStr);
			fclose($fp);
		}
	}

	/**
	 * 设置导出行数据
	 *
	 * @param array $columnList 行数据数组
	 * @param bool  $exportFlag 是否单行写入
	 */
	public function setXlsColumn($columnList, $exportFlag = true) {
		if (empty($columnList)) {
			return false;
		}

		$fp = fopen($this->tmpFile, 'a+');
		if ($exportFlag) {
			//转换导出数据
			foreach ($columnList as $key => $value) {
				$columnContent = $this->formatXlsColumn($value);
				fputs($fp, $columnContent);
			}
		//单行写入(不推荐)
		} else {
			$columnContent = $this->formatXlsColumn($columnList);
			fputs($fp, $columnContent);
		}

		fclose($fp);

		return true;
	}

	/**
	 * 导出数据文件
	 */
	public function exportXlsData() {
		header("Content-type:application/vnd.ms-excel");
		header("Content-Disposition:attachment;filename={$this->fileName}.xls");
		echo "<table width='100%' border='1'>";
		//分段获取大文件内容
		$offset = 0;
		do {
			$maxLen = self::MAX_LEN;
			$content = file_get_contents($this->tmpFile, NULL, NULL, $offset, $maxLen);
			if (!empty($content)) {
				echo $content;
				$offset += $maxLen;
			}
		} while (!empty($content));

		unlink($this->tmpFile);
		echo "</table>";
	}

	/**
	 * 格式化导出行数据
	 *
	 * @param array $columnValue 	行数据数组
	 * @param  boolean $isHeader 	是否为表头
	 * @return string $formatStr 	格式化字符串
	 */
	private function formatXlsColumn($columnValue, $isHeader = false) {
		if (empty($columnValue)) {
			return '';
		}
		$formatStr = "<tr>";
		foreach ($this->headerMap as $key => $value) {
			if ($isHeader) {
				$formatStr .= "<td bgcolor='{$this->styleOption['bgColor']}' style='font-size:{$this->styleOption['fontSize']}'>{$value['name']}</td>";
			} else {
				$styleType = in_array($value['type'], array_keys($this->styleMap)) ? $value['type'] : 'text';
				$style = "style='vnd.ms-excel.numberformat:{$this->styleMap[$styleType]};font-size:{$this->styleOption['fontSize']}'";
				$formatStr .= "<td align='left' {$style}>{$columnValue[$key]}</td>";
			}
		}
		$formatStr .= "</tr>";

		return $formatStr;
	}

	public function __destruct() {
		//删除临时存储文件
		if (file_exists($this->tmpFile)) {
			unlink($this->tmpFile);
		}
	}
}
