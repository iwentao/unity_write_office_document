using System;

public class Singleton<T> where T : class, new()
{
    public static T Instance
    {
        get { return Sub.instance; }
    }
    private class Sub
    {
        internal static readonly T instance = new T();
        static Sub() { }
    }
}
