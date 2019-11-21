package com.surpass.word;

import com.latextoword.Latex_Word;
import org.docx4j.XmlUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;

import javax.xml.bind.JAXBException;
import java.io.File;
import java.io.IOException;

/**
 * 将LaTeX值写入Word,且为可编辑的公式
 * <p>
 * Created by surpass.wei@gmail.com on 2019/11/21.
 */
public class WriteLaTeXFormula {
    private static ObjectFactory factory = Context.getWmlObjectFactory();

    public static void main(String[] args) throws JAXBException, Docx4JException, IOException {
        //  创建word数据包
        WordprocessingMLPackage mlPackage = WordprocessingMLPackage.createPackage();
        MainDocumentPart mainDoc = mlPackage.getMainDocumentPart();
        //  获取word公式对象
        String latex = "x = {-b \\pm \\sqrt{b^2-4ac} \\over 2a}";   //  随便输入一个标准的latex值
        Object wordFormulaObj = getWordFormulaObj(latex);
        //  创建word的段落
        P p = factory.createP();
        //  将公式对象添加到段落中
        p.getContent().add(wordFormulaObj);
        mainDoc.addObject(p);
        //  创建文件并写入数据
        File docx = File.createTempFile("word_", ".docx");
        mlPackage.save(docx);
    }

    /**
     * 获取LaTeX值对应的word公式对象
     *
     * @param latex LaTeX值
     * @return 可直接添加到文档中的公式对象
     * @throws JAXBException
     */
    private static Object getWordFormulaObj(String latex) throws JAXBException {
        String convertResult = Latex_Word.latexToWordAlreadyClean(latex);
        String requiredStr = " xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"" +
                " xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"";
        convertResult = new StringBuilder(convertResult).insert(8, requiredStr).toString();
        return XmlUtils.unmarshalString(convertResult);
    }
}
