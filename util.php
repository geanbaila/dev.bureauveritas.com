<?php
class Database{
	public function connect(){		
		$conn = mysqli_connect($GLOBALS["HOST"], $GLOBALS["USUARIO"], $GLOBALS["PASSWORD"], $GLOBALS["DATABASE"]); //or exit("Imposible conectar con la base de datos! ");
		echo $conn->connect_error;

		$conn->set_charset("utf8");
		return $conn;
	}

	public function disconnect($conn){
		$conn->close();
	}

	public function getResult($sql){
		if(DEBUG){echo "<pre>".$sql."</pre><hr>";}
		$conn = $this->connect();	
		$result = $conn->query($sql) or die("<br>Error al ejecutar:".$conn->error."<br><br>Detalle:<font color='red'>".$sql."</font><br><hr>");
		$data = array();
		$num_rows = $result->num_rows;
		if ($num_rows > 0) {
			$i=0;
		    while($row = $result->fetch_assoc()) {
		    	if(DEBUG){echo "<pre>";
		    	var_dump($row);
		    	echo "</pre>";}
		    	array_push($data, $row);
		    }
		}
		$this->disconnect($conn);
		return ["num_rows"=>$num_rows,"data"=>$data];
	}
}

	