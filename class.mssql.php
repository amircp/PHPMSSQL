<?php
/*****************************************************************+
 *
 *	MSSQL Server Database Class 
 *
 *	@Author 	Amir Canto
 *  @version	1.0
 *	@compatible	PHP 5 under Windows (Sql Server driver) & Linux (compiled with --mssql)
 *
 **/
if(!ini_get("safe_mode"))
{
	set_time_limit(0);
}

extract($_POST);
 
class MSSQL
{
	//atributos
	
	private $servidor,$loginServidor,$passServidor;
	private $baseDatos;
	private $handle = null;
	private $debug;
	
	
	
	public function getHandle()
	{
		return $this->handle;
	}
	//Constructor por Omisión
	
	/* Verify the connection state */
	public function isConnected()
	{
		if($this->handle != null) return true;
		else false;
	}
	public function __construct($host, $user,$pass,$db,$connect = false,$debug = false)
	{
	
		$this->servidor = $host;
		$this->loginServidor = $user;
		$this->passServidor = $pass;
		$this->baseDatos = $db;
		$this->debug = $debug;
		if($connect)
		{
			$this->connectMSSQL();
		}
	}	
	
	/* Execute Stored procedures */
	public function executeSProcedure($storedProcedure,$params)
	{
		if(function_exists("mssql_init"))
		{
			$stmt = mssql_init($storedProcedure,$this->handle);
			foreach($params as $keys => $values)
			{
				mssql_bind($stmt, "@".$keys,$values, SQLVARCHAR);
			}
			$arregloRespuesta = null;

			$resultado = mssql_execute($stmt);
			if ($resultado != 1)
			{
				while ($arreglo = mssql_fetch_array($resultado))
				{
					$arregloRespuesta[] = $arreglo;
				} 
			} 
			return $arregloRespuesta;
		}
		elseif (function_exists("sqlsrv_query"))
		{
			$param = null;
			$paramsSp = array();
			foreach($params as $key => $value)
			{
				$param .= "?,";
				$paramsSp[] = array($value, SQLSRV_PARAM_IN);
			}
			$param = substr($param,0, -1);
			$stmt = "{ call $storedProcedure ( $param ) }";

			$stmt_query = sqlsrv_query($this->handle,$stmt,$paramsSp) or exit(print_r(sqlsrv_errors(),true));
			if($stmt_query != false)
			{
				while($row = sqlsrv_fetch_array($stmt_query,SQLSRV_FETCH_ASSOC))
				{
					$arregloRespuesta[] = $row;
				}
				sqlsrv_free_stmt($stmt_query);
				return $arregloRespuesta;
			}
			return 0;
		
		}
		
	}
	
	
	
	/* Creates a connection to the mssql database server */
	public function connectMSSQL()
	{
		if(function_exists("mssql_connect"))
		{
		
			$this->handle = mssql_connect($this->servidor,$this->loginServidor, $this->passServidor) or exit("Error conectando a la base de datos");
			if($this->handle) mssql_select_db($this->baseDatos,$this->handle) or exit("No se encontr&oacute; la base de datos");
		}
		else if(function_exists("sqlsrv_connect"))
		{
			$serverName = $this->servidor;
			$connectionInfo = array("UID" => $this->loginServidor,"PWD" => $this->passServidor,"Database" => $this->baseDatos);
			$this->handle = sqlsrv_connect($serverName,$connectionInfo);
			if(!$this->handle) exit(print_r(sqlsrv_errors(),true));
		}
		else
		{
			exit("El driver de acceso no est&aacute; instalado.");
		}
		return 1;
	}	
	public function disconnectMSSQL()
	{
		if(function_exists("mssql_close"))
		{
			mssql_close($this->handle);
		} 
		else if(function_exists("sqlsrv_close"))
		{
			sqlsrv_close($this->handle);
		}
	}
	public function __destruct()
	{
		$this->disconnectMSSQL();
	}
	public function executeQuery($sentencia)
	{
	
		if($this->debug) echo "<pre>$sentencia</pre><br />";
		if(!$this->isConnected()) $this->connectMSSQL();
		$arregloRespuesta = null;

		if(function_exists("mssql_query"))
		{
			$resultado = mssql_query($sentencia) or exit("Error en la sentencia sql: ".$sentencia);
			
			while ($arreglo = mssql_fetch_array($resultado)){
				$arregloRespuesta[] = $arreglo;
			} //fin de while
			mssql_free_result($resultado);
			return $arregloRespuesta;
		}
		else if(function_exists("sqlsrv_query"))
		{

			$result = sqlsrv_query($this->handle,$sentencia) or exit(print_r(sqlsrv_errors(),true));
			if($result)
			{

				//$row = sqlsrv_fetch_array($result,SQLSRV_FETCH_ASSOC);
				
				while($row = sqlsrv_fetch_array($result,SQLSRV_FETCH_ASSOC))
				{
					$arregloRespuesta[] = $row;
				}
				sqlsrv_free_stmt($result);
				return $arregloRespuesta;
				//return $row;
			}
		}
	} //fin de function
		
	public function executeCommand($sentencia)
	{
		if($this->debug) echo "<pre>$sentencia</pre><br />";
		if(!$this->isConnected()) $this->connectMSSQL();
	
		if(function_exists("mssql_query"))
		{
			
			$resultado	= mssql_query($sentencia) or exit($sentencia); 
			mssql_free_result($resultado);
			return $resultado;
		}
		else if(function_exists("sqlsrv_query"))
		{
			$result = sqlsrv_query($sentencia) or exit(print_r(sqlsrv_errors(),true));
			if($result)
			{
				sqlsrv_free_stmt($result);
				return 1;
			}
		}
		return 0;
	}//fin de function
	
}//fin de class	
?>
