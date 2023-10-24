# 入门介绍
在格式为doc的word文件中书写模板指令，这些指令在运行时将被替换为实际传入的数据，最终输
出PDF文件。引擎采用数据和样式分离的概念，数据绑定替换的过程发生运行时，数据呈现样式由
模板文件控制。**即在制作模板的时候，为模板指令设定的样式如字体，颜色，对齐方式等将来会应 **
**用到实际的数据呈现中**。
**hello World Example**

如果你使用maven，将下面依赖添加到你的pom.xml文件中
```
<dependency> 
	<groupId>com.lianmed</groupId> 
	<artifactId>pdfgen</artifactId> 
	<version>1.6.8</version> 
</dependency>
```
在应用中添加如下类

```
public class Student {
    private String name;
    private Integer age;
    private List<Subject> subs;

    public Student(String name, Integer age) {
        super();
        this.name = name;
        this.age = age;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public List<Subject> getSubs() {
        return subs;
    }

    public void setSubs(List<Subject> subs) {
        this.subs = subs;
    }
}
```

```
public class Subject {
    private String name;
    private String score;

    public Subject(String name, String score) {
        super();
        this.name = name;
        this.score = score;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getScore() {
        return score;
    }

    public void setScore(String score) {
        this.score = score;
    }
}
```

在入口函数添加如下代码
```
Student data = new Student("张三", 22); 
//传入模板文件路径和数据 
WordHandler handler = new WordHandler("d:/11.doc", data ); 
//传入pdf输出路径 
handler.process("d:/11.pdf");
```

上述代码中传入的根数据为Student,所以用指令绑定Student中的属性时，直接引用属性名即可，
新建模板文件，添加内容如下，

![image.png](https://cdn.nlark.com/yuque/0/2023/png/21567582/1698134578108-e5f1f50b-7e04-4b15-b380-514b756569ec.png#averageHue=%23fdfdfd&clientId=uf974f6fb-d393-4&from=paste&height=169&id=uc6d9dead&originHeight=337&originWidth=747&originalType=binary&ratio=2&rotation=0&showTitle=false&size=11838&status=done&style=none&taskId=u78f941ec-8d2a-486e-adef-781c0302c15&title=&width=373.5)

运行程序即可得到目标PDF文件如下，可以看到指令${name}在模板中应用了粗体样式，目标PDF
中的“张三”也被应用了粗体

![image.png](https://cdn.nlark.com/yuque/0/2023/png/21567582/1698134624272-76c41e45-0df8-4ac4-95b2-463ce642d9c8.png#averageHue=%23fefefe&clientId=uf974f6fb-d393-4&from=paste&height=169&id=u8d524203&originHeight=337&originWidth=727&originalType=binary&ratio=2&rotation=0&showTitle=false&size=10608&status=done&style=none&taskId=u8ceee0db-34d4-4359-b8e7-528c7b4d421&title=&width=363.5)

# 模版语法

- **简单数据绑定**

$[O]{xxx} 其中O代表控制属性，xxx为数据绑定表达式，一切合法的ognl（对象图导航语言）表达
式都是合法的数据绑定表达式
当O不存在时，替换区域限制为模板中该模板指令占据的空间，当用实际数据替换该模板指令时，
超出该区域的部分会不显示。
当模板指令用下划线标识时，这时替换区域的空间不再是模板指令占据的空间，而是以该下划线开
始结束空间为替换区域空间。
**注意不要在一个下划线中写多个模板指令 **
当O存在时，即不限定替换区域，不管实际数据长度大小都能得到显示，但此时可能会对该行的布
局产生缩进。

- **表格中的集合遍历**

#[D]{xxx} 该指令必须写在表格中，其中D用于控制表格是否动态新增（不会动态减小）， xxx数据
绑定表达同样可以是一切合法的ongl表达式
当希望表格能根据运行时传入数据动态新增行时，只需在表格的任意一列的指令添加D控制属性
使用[idx]固定写法标识要遍历的属性，目前表格的遍历指令最多支持两层遍历，即xxx表达式中最
多包含两个[idx]标识
Example:
入口函数添加如下代码
```
Map<String, Object> data = new HashMap<String, Object>();
        List<Student> students = new ArrayList<Student>();
        Student s1 = new Student("张三", 22);
        List<Subject> sub1 = new ArrayList<Subject>();
        sub1.add(new Subject("语文", "80"));
        sub1.add(new Subject("数学", "67"));
        sub1.add(new Subject("英语", "35"));
        s1.setSubs(sub1);
        Student s2 = new Student("李四", 22);
        List<Subject> sub2 = new ArrayList<Subject>();
        sub2.add(new Subject("语文", "70"));
        s2.setSubs(sub2);
        Student s3 = new Student("王二", 22);
        students.add(s1);
        students.add(s2);
        students.add(s3);
        data.put("stus", students); 
        //传入模板输入流和数据 
        WordHandler handler = new WordHandler("d:/11.doc", data ); 
        //传入输出流接收生成的pdf 
        handler.process("d:/11.pdf");
```
使用如下图所示的模板

![image.png](https://cdn.nlark.com/yuque/0/2023/png/21567582/1698134923511-2f09f4de-6b82-4f89-82fb-7c2f028c256b.png#averageHue=%23faf9f9&clientId=uf974f6fb-d393-4&from=paste&height=180&id=u177dfff3&originHeight=360&originWidth=735&originalType=binary&ratio=2&rotation=0&showTitle=false&size=37702&status=done&style=none&taskId=uc8b4ffe9-a192-484e-af55-cd35bdfbd01&title=&width=367.5)

运行程序得到下图所示PDF

![image.png](https://cdn.nlark.com/yuque/0/2023/png/21567582/1698134953905-f3453bf7-9b64-447e-ba33-59c0753ee36c.png#averageHue=%23f9f9f9&clientId=uf974f6fb-d393-4&from=paste&height=184&id=ua46a1e10&originHeight=368&originWidth=732&originalType=binary&ratio=2&rotation=0&showTitle=false&size=36140&status=done&style=none&taskId=udc5f9f65-15d7-4ad5-b328-9e6e3f4a522&title=&width=366)   

如上所示，对集合的两级遍历会产生单元格合并

- **复选框**

@{xxx} 该指令一般位于批注中，同时该批注标识一个复选框，该复选框通过插入符号选项进行插
入，注意，xxx表达式必须是一个值为
布尔值的数据绑定表达式
如下是一个案例模板

![image.png](https://cdn.nlark.com/yuque/0/2023/png/21567582/1698135013488-2546e932-9907-4884-8890-e03bb67c13b0.png#averageHue=%23f9f7f6&clientId=uf974f6fb-d393-4&from=paste&height=35&id=u30e636af&originHeight=69&originWidth=718&originalType=binary&ratio=2&rotation=0&showTitle=false&size=17304&status=done&style=none&taskId=u8391ef52-94b9-4c3f-b35e-0f05ad73986&title=&width=359)

请参考ongl表达式写法，可以轻易将非布尔型属性组合成一个返回值为布尔类型的表达式

- **条件块**

<<if [xxx]>> data <</if>> 当xxx数据绑定表达式的值是true时，则data会被呈现在最终PDF
文件中，否则不会呈现。data区域可以嵌套使用
其他模板指令

- **插入图片**

1. 插入一个文本框，该文本框用于控制图片的大小和位置
2. 在文本框中写<<image [xxx]>>指令，注意xxx数据绑定表达式所绑定的必须是一张图片的
   InputStream实例

# 常见问题
转换pdf后字体样式与模版不一样
编辑模版使用到的字体，在程序运行的机器上必须要有，
