﻿using System.Reflection;
using System.Runtime.InteropServices;

// アセンブリに関する一般情報は以下の属性セットを通して制御されます。
// アセンブリに関連付けられている情報を変更するには、
// これらの属性値を変更してください。
[assembly: AssemblyTitle("BenryPPT")]
[assembly: AssemblyDescription("")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("")]
[assembly: AssemblyProduct("BenryPPT")]
[assembly: AssemblyCopyright("Copyright © 2020")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// ComVisible を false に設定すると、その型はこのアセンブリ内で COM コンポーネントには
// 見えなくなります。このアセンブリ内で COM から型にアクセスする必要がある場合は、
// その型の ComVisible 属性を true に設定してください。
[assembly: ComVisible(false)]

// このプロジェクトが COM に公開される場合、次の GUID が typelib の ID になります
[assembly: Guid("8498f378-ca3f-429a-8cd7-d47d9cfe9845")]

// アセンブリのバージョン情報は次の 4 つの値で構成されています:
//
//      メジャー バージョン
//      マイナー バージョン
//      ビルド番号
//      リビジョン
//
// すべての値を指定するか、以下のように '*' を使ってビルドおよびリビジョン番号を
// 既定値にすることができます:
// デフォルトではAssemblyVersionとAssemblyFileVersionがあるのですが、なぜか両方あると*が効きませんよね。
// AssemblyFileVersionを指定しなければ自動的なリビジョンが振られます。
// ただし、これだとファイルのプロパティから確認するときに製品バージョンにも同じ値が入ってしまいます。
// やはり製品バージョンは手動で値を振りたい。
// 製品バージョンはAssemblyInformationalVersionを使うと振れます。

[assembly: AssemblyVersion("1.7.0.*")]
// [assembly: AssemblyFileVersion("1.3.1.*")]
[assembly: AssemblyInformationalVersion("1.7.0")]
