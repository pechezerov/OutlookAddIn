﻿//------------------------------------------------------------------------------
// <auto-generated>
//     Этот код создан программой.
//     Исполняемая версия:4.0.30319.42000
//
//     Изменения в этом файле могут привести к неправильной работе и будут потеряны в случае
//     повторной генерации кода.
// </auto-generated>
//------------------------------------------------------------------------------

namespace QuoteAddIn.Properties {
    using System;
    
    
    /// <summary>
    ///   Класс ресурса со строгой типизацией для поиска локализованных строк и т.д.
    /// </summary>
    // Этот класс создан автоматически классом StronglyTypedResourceBuilder
    // с помощью такого средства, как ResGen или Visual Studio.
    // Чтобы добавить или удалить член, измените файл .ResX и снова запустите ResGen
    // с параметром /str или перестройте свой проект VS.
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "16.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    internal class Resources {
        
        private static global::System.Resources.ResourceManager resourceMan;
        
        private static global::System.Globalization.CultureInfo resourceCulture;
        
        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal Resources() {
        }
        
        /// <summary>
        ///   Возвращает кэшированный экземпляр ResourceManager, использованный этим классом.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Resources.ResourceManager ResourceManager {
            get {
                if (object.ReferenceEquals(resourceMan, null)) {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("QuoteAddIn.Properties.Resources", typeof(Resources).Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
        
        /// <summary>
        ///   Перезаписывает свойство CurrentUICulture текущего потока для всех
        ///   обращений к ресурсу с помощью этого класса ресурса со строгой типизацией.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Globalization.CultureInfo Culture {
            get {
                return resourceCulture;
            }
            set {
                resourceCulture = value;
            }
        }
        
        /// <summary>
        ///   Поиск локализованного ресурса типа System.Drawing.Bitmap.
        /// </summary>
        internal static System.Drawing.Bitmap Quote {
            get {
                object obj = ResourceManager.GetObject("Quote", resourceCulture);
                return ((System.Drawing.Bitmap)(obj));
            }
        }
        
        /// <summary>
        ///   Ищет локализованную строку, похожую на &lt;div class=WordSection1 style=&apos;font-family:Calibri&apos;&gt;
        ///&lt;p class=MsoNormal style=&apos;color:#A6A6A6;font-family:Calibri&apos;&gt;[%QUOTE_AUTHOR%] писал(а):&lt;o:p&gt;&lt;/o:p&gt;&lt;/p&gt;
        ///
        ///&lt;div class=WordSection1 style=&apos;background:#eaf5fa&apos;&gt;
        ///&lt;p class=MsoNormal style=&apos;margin-left:15pt;margin-top:5pt;margin-bottom:5pt;font-family:Calibri&apos;&gt;&lt;i&gt;
        ///&lt;span&gt;&lt;o:p&gt;[%QUOTE_TEXT%]&lt;/o:p&gt;&lt;/i&gt;&lt;/p&gt;
        ///&lt;p class=MsoNormal style=&apos;background:#eaf5fa&apos;&gt;&lt;/p&gt;
        ///&lt;/div&gt;
        ///
        ///&lt;/div&gt;.
        /// </summary>
        internal static string QuoteTemplate {
            get {
                return ResourceManager.GetString("QuoteTemplate", resourceCulture);
            }
        }
    }
}
