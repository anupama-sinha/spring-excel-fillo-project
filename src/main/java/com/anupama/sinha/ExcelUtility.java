package com.anupama.sinha;

import com.codoid.products.exception.FilloException;
import com.codoid.products.fillo.Connection;
import com.codoid.products.fillo.Fillo;
import com.codoid.products.fillo.Recordset;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RestController;

@RestController
public class ExcelUtility{
    @GetMapping("/upload")
    public String readExcel() throws FilloException{
        Fillo fillo=new Fillo();
        Connection connection=fillo.getConnection("C:\\Users\\anupa\\Documents\\Excel\\Test.xlsx");

        // Randomly reading data from Excel as String and then parsing to given datatype
        String strQuery="Select * from Sheet1 where ID=3";
        Recordset recordset=connection.executeQuery(strQuery);

        while(recordset.next()){
            String id = recordset.getField("ID");
            String name = recordset.getField("name");
            String dept = recordset.getField("department");
            String error = recordset.getField("error");

            System.out.println(id + " " + name + " " + dept + " " + error);
        }
        recordset.close();


        // Checking for certain condition in columns & updating the same
        String strQuery1="Update Sheet1 Set error='Blank name' where ID=3";
        connection.executeUpdate(strQuery1);

        // Deleting an entry from Excel based on Condition
        String strQuery2="Delete from Sheet1 where ID=2";
        connection.executeUpdate(strQuery2);

        connection.close();

        return "done";
    }

}
