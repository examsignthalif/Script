using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace Script
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public static bool AcceptOnlyIntegers(object sender)
        {
            TextBox textBox = sender as TextBox;

            if (textBox.Text.Length <= 0)
                return false;

            int selectionStart = textBox.SelectionStart;
            string newText = string.Empty;
            int count = 0;
            foreach (char c in textBox.Text.ToCharArray())
            {
                if (char.IsDigit(c) || char.IsControl(c) || (c == '.' && count == 0))
                {
                    newText += c;
                    if (c == '.')
                        count += 1;
                }
            }
            textBox.Text = newText;
            textBox.SelectionStart = selectionStart <= textBox.Text.Length ? selectionStart : textBox.Text.Length;
            return true;
        }
    }
}
