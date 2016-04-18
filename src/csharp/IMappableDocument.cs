namespace Office
{
    using System;

    public interface IMappableDocument : IDisposable
    {
        void MapNamedRangesTo<T>(T instance);

        void Commit(string outputPath);
    }
}