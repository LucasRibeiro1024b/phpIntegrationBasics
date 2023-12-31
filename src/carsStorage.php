<?php

namespace phpsheet;

class CarsStorage {
	protected $items = [];

	public function __construct() {
		$fp = fopen('docs/cars.csv', 'r');
		while($line = fgetcsv($fp)) {
			$this->items[$line[0]] = 
				["url" => "http://localhost:8080/api.php/cars/" . $line[0], 
				"name" => $line[1], "link" => $line[2]];
		}
		fclose($fp);
	}

	public function create($data) {
		if(isset($data['name']) && isset($data['link'])) {
			$id = $this->generateID();
			$this->items[$id] = [
				"url" => $this->makeUrlFromIndex($id),
				"name" => $data['name'],
				"link" => $data['link']
				];
			$item = $this->getOne($id);
			return $item;
		}
		throw new \UnexpectedValueException("Could not create item");
	}

	protected function generateID() {
		// even better, use ramsey/uuid
		return substr(sha1(microtime()),0,5);
	}

	public function getAll() {
		return array_values($this->items);
	}

	public function getOne($index) {
		if(isset($this->items[$index])) {
			return $this->items[$index];
		}
		// we didn't find it
		throw new \UnexpectedValueException("Item not found");
	}

	protected function makeUrlFromIndex($index)
	{
		return "http://localhost:8080/api.php/cars/" . $index;
	}

	public function remove($index) {
		if(isset($this->items[$index])) {
			unset($this->items[$index]);
		}
		return true;
	}

	public function save() {
		$fp = fopen('docs/cars.csv', 'w');
		foreach($this->items as $id => $item) {
			$line = [$id, $item['name'], $item['link']];
			fputcsv($fp, $line);
		}
		fclose($fp);
	}

	public function update($id, $data) {
		if(isset($data['name']) && isset($data['link'])) {
			$this->items[$id] = [
				"url" => $this->makeUrlFromIndex($id),
				"name" => $data['name'],
				"link" => $data['link']
				];
			$item = $this->getOne($id);
			return $item;
		}
		throw new \UnexpectedValueException("Could not create item");
	}

}