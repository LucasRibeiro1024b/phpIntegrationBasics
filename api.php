<?php

// require("ItemStorage.php");
include 'vendor/autoload.php';

use phpsheet\ItemStorage;

set_exception_handler(function ($e) {
	$code = $e->getCode() ?: 400;
	header("Content-Type: application/json", false, $code);
	echo json_encode(["error" => $e->getMessage()]);
	exit;
});

// assume JSON, handle requests by verb and path
$verb = $_SERVER['REQUEST_METHOD'];
$url_pieces = explode('/', $_SERVER['PATH_INFO']);
$class = $url_pieces[1];
$className = 'phpsheet\\'.$url_pieces[1].'Storage';
$storage = new $className();

// catch this here, we don't support many routes yet
if(!class_exists($className)) {
	throw new Exception('Unknown endpoint', 404);
}

switch($verb) {
	case 'GET':
		if(isset($url_pieces[2])) {
			try {
				$data = $storage->getOne($url_pieces[2]);
			} catch (UnexpectedValueException $e) {
				throw new Exception("Resource does not exist", 404);
			}
		} else {
			$data = $storage->getAll();
		}
		break;
	// two cases so similar we'll just share code
	case 'POST':
	case 'PUT':
		// read the JSON
		$params = json_decode(file_get_contents("php://input"), true);

		if(!$params) {
			throw new Exception("Data missing or invalid");
		}
		if($verb == 'PUT') {
			$id = $url_pieces[2];
			$item = $storage->update($id, $params);
			$status = 204;
		} else {
			$item = $storage->create($params);
			$status = 201;
		}
		$storage->save();

		// send header, avoid output handler
		header("Location: " . $item['url'], false, $status);
		exit;
		// break;
	case 'DELETE':
		$id = $url_pieces[2];
		$storage->remove($id);
		$storage->save();
		header("Location: http://localhost:8080/".$class, false, 204);
		exit;
		// break;
	default: 
		throw new Exception('Method Not Supported', 405);
}

header("Content-Type: application/json");
echo json_encode($data);