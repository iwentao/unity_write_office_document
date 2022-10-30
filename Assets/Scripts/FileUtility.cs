using System;
using System.IO;
using UnityEngine;

public class FileUtility
{
    public enum ImageType
    {
        Null,
        Png,
        Jpg,
        Gif,
        Bmp
    }

    /// <summary>
    /// 获取图片格式
    /// </summary>
    private static ImageType GetImageType(byte[] bytes)
    {
        byte[] header = new byte[8];
        Array.Copy(bytes, header, header.Length);
        ImageType type = ImageType.Null;
        //读取图片文件头8个字节
        //Png图片 8字节：89 50 4E 47 0D 0A 1A 0A   =  [1]:P[2]:N[3]:G
        if (header[0] == 0x89 && header[1] == 0x50 && header[2] == 0x4E && header[3] == 0x47 &&
            header[4] == 0x0D && header[5] == 0x0A && header[6] == 0x1A && header[7] == 0x0A)
        {
            type = ImageType.Png;
        }
        //Jpg图片 2字节：FF D8
        else if (header[0] == 0xFF && header[1] == 0xD8)
        {
            type = ImageType.Jpg;
        }
        //Gif图片 6字节：47 49 46 38 39|37 61   =   GIF897a
        else if (header[0] == 0x47 && header[1] == 0x49 && header[2] == 0x46 && header[3] == 0x38 &&
            (header[4] == 0x39 || header[4] == 0x37) && header[5] == 0x61)
        {
            type = ImageType.Gif;
        }
        //Bmp图片 2字节：42 4D
        else if (header[0] == 0x42 && header[1] == 0x4D)
        {
            type = ImageType.Bmp;
        }
        return type;
    }

    private static byte[] _header = null;
    private static byte[] _buffer = null;

    public static void FileInfo(string filePath, out byte[] bytes, out Vector2 size)
    {
        size = Vector2.zero;
        FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
        stream.Seek(0, SeekOrigin.Begin);
        bytes = new byte[stream.Length];
        stream.Read(bytes, 0, (int)stream.Length);

        ImageType imageType = GetImageType(bytes);
        switch (imageType)
        {
            case ImageType.Png:
                {
                    stream.Seek(0, SeekOrigin.Begin);
                    _header = new byte[8];
                    stream.Read(_header, 0, 8);
                    stream.Seek(8, SeekOrigin.Current);

                    _buffer = new byte[8];
                    stream.Read(_buffer, 0, _buffer.Length);

                    Array.Reverse(_buffer, 0, 4);
                    Array.Reverse(_buffer, 4, 4);

                    size.x = BitConverter.ToInt32(_buffer, 0);
                    size.y = BitConverter.ToInt32(_buffer, 4);
                }
                break;
            case ImageType.Jpg:
                {
                    stream.Seek(0, SeekOrigin.Begin);
                    _header = new byte[2];
                    stream.Read(_header, 0, 2);
                    //段类型
                    int type = -1;
                    int ff = -1;
                    //记录当前读取的位置
                    long ps = 0;
                    //逐个遍历所以段，查找SOFO段
                    do
                    {
                        do
                        {
                            //每个新段的开始标识为oxff，查找下一个新段
                            ff = stream.ReadByte();
                            if (ff < 0) //文件结束
                            {
                                return;
                            }
                        } while (ff != 0xff);

                        do
                        {
                            //段与段之间有一个或多个oxff间隔，跳过这些oxff之后的字节为段标识
                            type = stream.ReadByte();
                        } while (type == 0xff);

                        //记录当前位置
                        ps = stream.Position;
                        switch (type)
                        {
                            case 0x00:
                            case 0x01:
                            case 0xD0:
                            case 0xD1:
                            case 0xD2:
                            case 0xD3:
                            case 0xD4:
                            case 0xD5:
                            case 0xD6:
                            case 0xD7:
                                break;
                            case 0xc0: //SOF0段（图像基本信息）
                            case 0xc2: //JFIF格式的 SOF0段
                                {
                                    //找到SOFO段，解析宽度和高度信息
                                    //跳过2个自己长度信息和1个字节的精度信息
                                    stream.Seek(3, SeekOrigin.Current);

                                    //高度 占2字节 低位高位互换
                                    size.y = stream.ReadByte() * 256;
                                    size.y += stream.ReadByte();
                                    //宽度 占2字节 低位高位互换
                                    size.x = stream.ReadByte() * 256;
                                    size.x += stream.ReadByte();
                                    return;
                                }
                            default: //别的段都跳过
                                        //获取段长度，直接跳过
                                ps = stream.ReadByte() * 256;
                                ps = stream.Position + ps + stream.ReadByte() - 2;
                                break;
                        }
                        if (ps + 1 >= stream.Length) //文件结束
                        {
                            return;
                        }
                        stream.Position = ps; //移动指针
                    } while (type != 0xda); // 扫描行开始
                }
                break;
            case ImageType.Gif:
                {
                    stream.Seek(0, SeekOrigin.Begin);
                    _header = new byte[6];
                    stream.Read(_header, 0, 6);

                    _buffer = new byte[4];
                    stream.Read(_buffer, 0, _buffer.Length);

                    size.x = BitConverter.ToInt16(_buffer, 0);
                    size.y = BitConverter.ToInt16(_buffer, 2);
                }
                break;
            case ImageType.Bmp:
                {
                    stream.Seek(0, SeekOrigin.Begin);
                    _header = new byte[2];
                    stream.Read(_header, 0, 2);
                    //跳过16个字节
                    stream.Seek(16, SeekOrigin.Current);
                    //bmp图片的宽度信息保存在第 18-21位 4字节
                    //bmp图片的高度度信息保存在第 22-25位 4字节
                    _buffer = new byte[8];
                    stream.Read(_buffer, 0, _buffer.Length);

                    size.x = BitConverter.ToInt32(_buffer, 0);
                    size.y = BitConverter.ToInt32(_buffer, 4);
                }
                break;
            default:
                break;
        }

        stream.Close();
        stream.Dispose();
    }

}