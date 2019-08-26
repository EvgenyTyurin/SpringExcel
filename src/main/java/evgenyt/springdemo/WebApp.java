package evgenyt.springdemo;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.servlet.ModelAndView;

import java.util.ArrayList;
import java.util.List;

/**
 * WebApp servlet controller
 */

@Controller
public class WebApp {
    // Show home.jsp
    @RequestMapping(value = "/", method = RequestMethod.GET)
    public String getHome() {
        return "/";
    }
    // Make data and provide excel file
    @RequestMapping(value = "/report", method=RequestMethod.GET)
    public ModelAndView getExcel(){
        // Make data
        List<String> list = new ArrayList<>();
        list.add("test1");
        list.add("test2");
        // Make Excel file
        return new ModelAndView(new ExcelReport(), "testList", list);
    }
}
