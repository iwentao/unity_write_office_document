using System.Collections;
using NPOI.XWPF.UserModel;
using System.Collections.Generic;
using UnityEngine;

public class TestClient : MonoBehaviour
{
    void Start()
    {
        Debug.Log("start to write doc");
        // WordManager.Instance.TargetPath = Application.streamingAssetsPath + "/test.docx";
        // WordManager.Instance.Write2();

        // WordManager.Instance.TargetPath = Application.streamingAssetsPath + "/test2.docx";
        // WordManager.Instance.Write3();

        WordManager.Instance.TargetPath = Application.streamingAssetsPath + "/test7.docx";
        string content = "日前，在四川自贡凤鸣通航机场，一架代号为“双尾蝎D”的大型四发无人机自主驶入跑道并顺利升空，飞行18分钟后平稳着陆，全程无故障，首飞圆满成功。\n该型无人机是我国自主研发，拥有完全知识产权的大型四发无人机，具备了更大的装载空间、更重的装载能力、更强的用电支持和更高的系统可靠性、飞行安全性，可携带更多高性能的任务载荷，执行货运物流、航空播撒、任务载荷使用等支援保障任务。\n此次，大型四发无人机实现首飞，进一步丰富了商用大型无人机的应用场景，也标志着我国大型无人机产业已具备针对不同市场需求快速研发响应的能力。“双尾蝎D”成功首飞后将转场珠海，亮相即将举办的第十四届中国航展。\n共青团中央共青团工作、活动信息和青年关注的热点信息";
        WordManager.Instance.AddTitle("祝贺！国产大型四发无人机成功首飞！", WordManager.Font_Color_Red);
        WordManager.Instance.WriteArticle(content, '\n', false);
        string[] table_content = {"row0", "许昌市魏都区和西华县等五地发布致在富士康工作人员的公开信", "row2 定住所，提供持", "富士康科技集团发布声明表示，网络流传的“郑州园区约两万人确诊”为严重不实信息。富士康方面对第一财经记者表示"};
        WordManager.Instance.AddTable(2, 2, new float[]{6, 8}, table_content);
        string image_folder = Application.dataPath + "/../test_images";
        // string image_name = "test3.png";
        // WordManager.Instance.AddImage(image_folder, image_name, 0.6f);
        string[] image_names = {"test3.png", "test.jpg"};
        WordManager.Instance.AddImages(image_folder, image_names, new float[]{0.6f, 0.5f});

        WordManager.Instance.Save();

        WordManager.Instance.RegenDocument();
        WordManager.Instance.TargetPath = Application.streamingAssetsPath + "/test8.docx";
        string content2 = "爱我中华，我爱中华\tHello, world!\t This is a wonderful world";
        WordManager.Instance.AddTitle("我和我的祖国", WordManager.Font_Color_Red, 18, true);
        WordManager.Instance.WriteArticle(content2, '\t', false);
        WordManager.Instance.Save();

        WordManager.Instance.SourcePath = Application.streamingAssetsPath + "/report_template.docx";
        WordManager.Instance.TargetPath = Application.streamingAssetsPath + "/test9.docx";
        WordManager.Instance.ReplaceTable = new Dictionary<string, string>(){
            {"{%template_title%}", "系统运行情况分析"},
            {"{%template_date%}", System.DateTime.Now.ToString()}
        };
        WordManager.Instance.ReplaceWord(false);
        WordManager.Instance.AddParagraph("我的测试段落", ParagraphAlignment.LEFT);

        WordManager.Instance.Save();
    }

    void Update()
    {
    }
}
