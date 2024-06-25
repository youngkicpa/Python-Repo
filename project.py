import zipfile
import os

# 프로젝트 디렉토리 구조와 파일 생성
project_dir = "/CardProject"
os.makedirs(project_dir, exist_ok=True)

# 파일 내용 작성
files = {
    "App.xaml": """
<Application x:Class="CardProject.App"
             xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
             xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
             StartupUri="MainWindow.xaml">
    <Application.Resources>
         
    </Application.Resources>
</Application>
""",
    "App.xaml.cs": """
using System.Windows;

namespace CardProject
{
    public partial class App : Application
    {
    }
}
""",
    "MainWindow.xaml": """
<Window x:Class="CardProject.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="MainWindow" Height="450" Width="800">
    <Grid Name="CardGrid">
        
    </Grid>
</Window>
""",
    "MainWindow.xaml.cs": """
using System.Windows;
using System.Windows.Controls;

namespace CardProject
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            CreateCards();
        }

        private void CreateCards()
        {
            for (int i = 1; i <= 10; i++)
            {
                Button cardButton = new Button
                {
                    Content = $"Card {i}",
                    Name = $"Card{i}",
                    Margin = new Thickness(10),
                    Width = 100,
                    Height = 50
                };
                cardButton.Click += CardButton_Click;
                CardGrid.Children.Add(cardButton);
            }
        }

        private void CardButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button)
            {
                MessageBox.Show($"Clicked: {button.Content}");
            }
        }
    }
}
""",
    "CardProject.csproj": """
<Project Sdk="Microsoft.NET.Sdk.WindowsDesktop">

  <PropertyGroup>
    <OutputType>WinExe</OutputType>
    <TargetFramework>net8.0</TargetFramework>
    <UseWPF>true</UseWPF>
  </PropertyGroup>

</Project>
"""
}

# 파일 생성 및 내용 작성
for filename, content in files.items():
    with open(os.path.join(project_dir, filename), 'w') as file:
        file.write(content.strip())

# 압축 파일 생성
zip_filename = "/CardProject/CardProject.zip"
with zipfile.ZipFile(zip_filename, 'w') as zf:
    for root, _, files in os.walk(project_dir):
        for file in files:
            file_path = os.path.join(root, file)
            zf.write(file_path, os.path.relpath(file_path, project_dir))

zip_filename
