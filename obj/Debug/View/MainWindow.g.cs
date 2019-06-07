﻿#pragma checksum "..\..\..\View\MainWindow.xaml" "{ff1816ec-aa5e-4d10-87f7-6f4963833460}" "29C560D7330A1305083646D7CC33CBD7455C3E98"
//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

using Midi_Analyzer;
using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Media.TextFormatting;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Shell;


namespace Midi_Analyzer {
    
    
    /// <summary>
    /// MainWindow
    /// </summary>
    public partial class MainWindow : System.Windows.Window, System.Windows.Markup.IComponentConnector, System.Windows.Markup.IStyleConnector {
        
        
        #line 11 "..\..\..\View\MainWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TabControl tabController;
        
        #line default
        #line hidden
        
        
        #line 20 "..\..\..\View\MainWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox excerptBox;
        
        #line default
        #line hidden
        
        
        #line 27 "..\..\..\View\MainWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox modelBox;
        
        #line default
        #line hidden
        
        
        #line 34 "..\..\..\View\MainWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.ListBox sourcePath;
        
        #line default
        #line hidden
        
        
        #line 36 "..\..\..\View\MainWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox destinationPath;
        
        #line default
        #line hidden
        
        
        #line 45 "..\..\..\View\MainWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.RadioButton midiButton;
        
        #line default
        #line hidden
        
        
        #line 46 "..\..\..\View\MainWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.RadioButton csvButton;
        
        #line default
        #line hidden
        
        
        #line 51 "..\..\..\View\MainWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TabItem errorDetection;
        
        #line default
        #line hidden
        
        
        #line 53 "..\..\..\View\MainWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.ListBox xlsFileList;
        
        #line default
        #line hidden
        
        private bool _contentLoaded;
        
        /// <summary>
        /// InitializeComponent
        /// </summary>
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
        public void InitializeComponent() {
            if (_contentLoaded) {
                return;
            }
            _contentLoaded = true;
            System.Uri resourceLocater = new System.Uri("/Midi Analyzer;component/view/mainwindow.xaml", System.UriKind.Relative);
            
            #line 1 "..\..\..\View\MainWindow.xaml"
            System.Windows.Application.LoadComponent(this, resourceLocater);
            
            #line default
            #line hidden
        }
        
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
        [System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes")]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity")]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1800:DoNotCastUnnecessarily")]
        void System.Windows.Markup.IComponentConnector.Connect(int connectionId, object target) {
            switch (connectionId)
            {
            case 1:
            this.tabController = ((System.Windows.Controls.TabControl)(target));
            return;
            case 2:
            this.excerptBox = ((System.Windows.Controls.TextBox)(target));
            return;
            case 3:
            
            #line 24 "..\..\..\View\MainWindow.xaml"
            ((System.Windows.Controls.Button)(target)).Click += new System.Windows.RoutedEventHandler(this.BrowseForCSV);
            
            #line default
            #line hidden
            return;
            case 4:
            this.modelBox = ((System.Windows.Controls.TextBox)(target));
            return;
            case 5:
            
            #line 31 "..\..\..\View\MainWindow.xaml"
            ((System.Windows.Controls.Button)(target)).Click += new System.Windows.RoutedEventHandler(this.BrowseForMidi);
            
            #line default
            #line hidden
            return;
            case 6:
            this.sourcePath = ((System.Windows.Controls.ListBox)(target));
            
            #line 34 "..\..\..\View\MainWindow.xaml"
            this.sourcePath.KeyDown += new System.Windows.Input.KeyEventHandler(this.DeleteItem);
            
            #line default
            #line hidden
            return;
            case 7:
            this.destinationPath = ((System.Windows.Controls.TextBox)(target));
            return;
            case 8:
            
            #line 39 "..\..\..\View\MainWindow.xaml"
            ((System.Windows.Controls.Button)(target)).Click += new System.Windows.RoutedEventHandler(this.PopulateSourceListbox);
            
            #line default
            #line hidden
            return;
            case 9:
            
            #line 41 "..\..\..\View\MainWindow.xaml"
            ((System.Windows.Controls.Button)(target)).Click += new System.Windows.RoutedEventHandler(this.BrowseForFolder);
            
            #line default
            #line hidden
            return;
            case 10:
            
            #line 43 "..\..\..\View\MainWindow.xaml"
            ((System.Windows.Controls.Button)(target)).Click += new System.Windows.RoutedEventHandler(this.ConvertFile);
            
            #line default
            #line hidden
            return;
            case 11:
            this.midiButton = ((System.Windows.Controls.RadioButton)(target));
            
            #line 45 "..\..\..\View\MainWindow.xaml"
            this.midiButton.Checked += new System.Windows.RoutedEventHandler(this.buttonCheckChange);
            
            #line default
            #line hidden
            return;
            case 12:
            this.csvButton = ((System.Windows.Controls.RadioButton)(target));
            
            #line 46 "..\..\..\View\MainWindow.xaml"
            this.csvButton.Checked += new System.Windows.RoutedEventHandler(this.buttonCheckChange);
            
            #line default
            #line hidden
            return;
            case 13:
            
            #line 48 "..\..\..\View\MainWindow.xaml"
            ((System.Windows.Controls.Button)(target)).Click += new System.Windows.RoutedEventHandler(this.AnalyzeFile);
            
            #line default
            #line hidden
            return;
            case 14:
            this.errorDetection = ((System.Windows.Controls.TabItem)(target));
            return;
            case 15:
            this.xlsFileList = ((System.Windows.Controls.ListBox)(target));
            return;
            }
            this._contentLoaded = true;
        }
        
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
        [System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes")]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1800:DoNotCastUnnecessarily")]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity")]
        void System.Windows.Markup.IStyleConnector.Connect(int connectionId, object target) {
            System.Windows.EventSetter eventSetter;
            switch (connectionId)
            {
            case 16:
            eventSetter = new System.Windows.EventSetter();
            eventSetter.Event = System.Windows.Controls.Control.MouseDoubleClickEvent;
            
            #line 56 "..\..\..\View\MainWindow.xaml"
            eventSetter.Handler = new System.Windows.Input.MouseButtonEventHandler(this.OpenFile);
            
            #line default
            #line hidden
            ((System.Windows.Style)(target)).Setters.Add(eventSetter);
            break;
            }
        }
    }
}

