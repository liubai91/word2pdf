package org.word2pdf.processor;

import com.aspose.words.Comment;
import com.aspose.words.CommentCollection;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.lianmed.pdf.WordHandleContext;
import com.lianmed.pdf.WordHandler;
import com.lianmed.processor.Processor;
import groovy.lang.Binding;
import groovy.lang.GroovyShell;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.BeanUtils;

import java.beans.PropertyDescriptor;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

public class GroovyProcessor implements Processor {

    private final Logger log = LoggerFactory.getLogger(GroovyProcessor.class);

    //脚本文本
    //脚本文件名
    private static final Pattern IMG_DIRECTIVE_PATTERN = Pattern.compile("(<<image.*?\\[)(.*)(\\].*?>>)");


    @Override
    public void process(WordHandleContext context) throws Exception {



        List<String> scriptfiles = new ArrayList<>();
        List<String> scripts = new ArrayList<>();



        //解析指令
        NodeCollection childs = context.getDoc().getChildNodes(NodeType.COMMENT, true);
        for (Object child : childs) {
            Comment comment = (Comment) child;

            if(comment.getText().startsWith("groovy")) {
                String text = comment.getText();
                /*if(text.startsWith("file:")) {
                    scriptfiles.add(text.substring(5).trim());
                }else {
                    scripts.add(text);
                }*/
                scripts.add(text.substring(6));
                comment.remove();
            }
        }


        //执行groovy脚本,groovy脚本主要是data中插入新属性
        Map<String,Object> derivation = new HashMap<>();
        Binding binding = new Binding();
        binding.setVariable("input",context.getData());
        binding.setVariable("output", derivation);
        GroovyShell shell = new GroovyShell(WordHandler.class.getClassLoader(),binding);
        for (String script : scripts) {
            try {
                shell.evaluate(script);
            } catch (Exception e) {
                log.info("脚本执行错误",e);
            }
        }

        //转换data
        Map result = convertData(context.getData());
        if(result == null) {
            result = new HashMap();
        }
        result.putAll(derivation);
        context.setData(result);


    }

    private Map convertData(Object data) {
        if(data ==null || data instanceof Map) {
            return (Map) data;
        } else {
            Map<String,Object> result = new HashMap<>();
            try {
                Class<?> dataClass = data.getClass();
                PropertyDescriptor[] propertyDescriptors = BeanUtils.getPropertyDescriptors(dataClass);
                for (PropertyDescriptor propertyDescriptor : propertyDescriptors) {
                    if (propertyDescriptor.getReadMethod()!=null) {
                        result.put(propertyDescriptor.getName(),propertyDescriptor.getReadMethod().invoke(data));
                    }
                }

            }catch (Exception e) {

            }
            return result;
        }
    }
}
