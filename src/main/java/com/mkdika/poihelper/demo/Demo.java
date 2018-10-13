package com.mkdika.poihelper.demo;

import com.mkdika.jeneric.function.DateFun;
import com.mkdika.poihelper.helper.CollectionToPlainExcel;

import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.List;

/**
 * Demo and usage example.
 *
 * @author Maikel Chandika <mkdika@gmail.com>
 * @version 1.0.0
 * @since 11th Oct 2018
 */
public class Demo {

    private static List<Person> data1;
    private static List<String> data2;

    static {
        // init data1
        data1 = new ArrayList<>();
        data1.add(new Person(1, "Maikel", DateFun.of(1990, 3, 19), 99.5));
        data1.add(new Person(2, "Budi", DateFun.of(1995, 4, 1), 9.5));
        data1.add(new Person(3, "Johan", DateFun.of(1980, 5, 2), 10d));
        data1.add(new Person(4, "Andy", DateFun.of(1985, 6, 4), 55.43));
        data1.add(new Person()); // will result a empty line :)
        data1.add(null); // will not generated
        data1.add(new Person(5, "Irwan", DateFun.of(1997, 7, 10), 70.0));
        data1.add(new Person(6, null, DateFun.of(1991, 10, 11), 60.0));

        // init data2
        data2 = new ArrayList<>();
        data2.add("John");
        data2.add("Felix");
        data2.add("Max");
    }

    public static void main(String[] args) throws ClassNotFoundException, IOException, IllegalAccessException, NoSuchMethodException, InvocationTargetException {
        CollectionToPlainExcel coe = new CollectionToPlainExcel();
        coe.setData(data1);

        // write to file example
        coe.writeToExcel("c:\\test\\test1.xlsx");

        // to array of byte example
        byte[] result = coe.writeToExcel();
        FileOutputStream stream = new FileOutputStream("c:\\test\\test2.xlsx");
        try {
            stream.write(result);
        } finally {
            stream.close();
        }

        // test simple type data
        coe.setData(data2);
        coe.writeToExcel("c:\\test\\test3.xlsx");
    }
}
