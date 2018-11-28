using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Extended.Npoi
{
#if NET40
    /* .Net 4.5 新增接口
     * 
     * 官方推荐异步方法重载:
     * public Task MethodNameAsync(...);
     * public Task MethodNameAsync(..., CancellationToken cancellationToken);
     * public Task MethodNameAsync(..., IProgress<T> progress); 
     * public Task MethodNameAsync(..., CancellationToken cancellationToken, IProgress<T> progress);
     */
    public interface IProgress<in T>
    {
        void Report(T value);
    }
#endif
}
