package se.niteco.qa;

import se.niteco.qa.model.ExampleModel;
import se.niteco.qa.utils.ExcelDeserializor;

import java.util.List;

/**
 * Created by khoi.nguyen on 7/6/2016.
 */
public class Example {
	public static void main(String[] args) {
		List object = new ExcelDeserializor().convert("D:\\test.xlsx", ExampleModel.class);
	}
}