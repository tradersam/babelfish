using System;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;

namespace BabbleFish
{
    public class Glossary
    {
        public Glossary( string name )
        {
            _name = name;
            entries = new Dictionary<string, string>();
            errors = new List<string>();
        }

        string _name;
        public string Name { get { return _name; } }
        //todo language support

        Dictionary<string, string> entries;
        public Dictionary<string, string> Entries { get { return entries; } }

        List<string> errors;
        public List<string> ErrorMessages{ get { return errors; } }
        const string errorFormat = "string id {0} duplicated. Was: \"{1}\" Now: \"{2}\"";

        public void AddEntry( KeyValuePair<string, string> stringID_locString )
        {
            //This method is cheeky, maybe cleanup?
            AddEntry( stringID_locString.Key, stringID_locString.Value );
        }

        public void AddEntry( string stringID, string locString )
        {
            string lowerKey = stringID.ToLower();

            string oldLocString;
            if( entries.TryGetValue( lowerKey, out oldLocString ) )
            {
                TrackError( lowerKey, locString, oldLocString );
            }
            entries[lowerKey] = locString;
        }

        private void TrackError( string stringID, string locString, string oldLocString )
        {
            errors.Add( string.Format( errorFormat, stringID, locString, oldLocString ) );
        }
    }

    public class SourceFile
    {
        ISheet workSheet;

        string _filePath;
        string FilePath { get { return _filePath; } }

        const string stringID = "stringID";
        const string english = "EN";    //ISO 639-2 https://www.loc.gov/standards/iso639-2/

        int stringIDCol = -1;
        int englishCol = -1;    //todo, eventually make a set of objects that know their language and offset in the doc

        public SourceFile( string filePath )
        {
            _filePath = filePath;

            OpenWorkbook();
        }

        private void OpenWorkbook()
        {
            try
            {
                using( FileStream file = new FileStream( _filePath, FileMode.Open, FileAccess.Read ) )
                {
                    IWorkbook workbook = new XSSFWorkbook( file );

                    if( workbook.NumberOfSheets > 1 )
                    {
                        Console.WriteLine( "Warning! More than one sheet present in this doc. Only the first sheet will be processed. Please delete the remaining sheets" );
                    }

                    workSheet = workbook.GetSheetAt( 0 );
                }
            }
            catch( Exception e )
            {
                Program.ErrorExit( e.ToString() );
            }

            ParseHeader();

            rowOffset = 1;  //Skip row 0 since the header lives there
        }

        private void ParseHeader()
        {
            var headerRow = workSheet.GetRow( 0 );
            for( int i = headerRow.FirstCellNum; i < headerRow.LastCellNum; ++i )
            {
                string str = headerRow.GetCell( i ).ToString();

                if( string.Compare( str, stringID, true ) == 0 )
                {
                    if( stringIDCol != -1 )
                    {
                        Program.ErrorExit( "StringID column found twice. Remove one and try again." );
                    }

                    stringIDCol = i;
                }
                else if( string.Compare( str, english, true ) == 0 )
                {
                    if( englishCol != -1 )
                    {
                        //Todo rework to support more languages and be more generic/flexible
                        Program.ErrorExit( "EN column found twice. Remove one and try again." );
                    }

                    englishCol = i;
                }
            }

            //validate all things
            if( stringIDCol == -1 )
            {
                Program.ErrorExit( "Unable to find StringID column, add it to row 1 and retry" );
            }

            if( englishCol == -1 )
            {
                Program.ErrorExit( "Unable to find EN column, add it to row 1 and retry" );
            }
        }

        int rowOffset;

        public bool HasData()
        {
            return workSheet != null && rowOffset <= workSheet.LastRowNum;
        }

        private void MoveNext()
        {
            rowOffset++;
        }

        public KeyValuePair<string, string> ReadEntry()
        {
            var row = workSheet.GetRow( rowOffset );

            string key = row.GetCell( stringIDCol ).ToString();
            string value = row.GetCell( englishCol ).ToString();

            MoveNext();

            return new KeyValuePair<string, string>( key, value );
        }

    }

    class Program
    {
        static public bool WarnOnDuplicates = false;
        //static public bool AllowEmptyCells = false;   //TODO, warn if cell has nothing in it

        static void Main( string[] args )
        {
            string filePath = "";

            if( args.Length == 0 )
            {
                PrintInfo();


                filePath = "../../../ExampleDictionary/example.xlsx";
                Console.WriteLine( "Ignore that message for now, we're going to use an example dictionary" );
            }
            else if( args.Length == 1 )
            {
                filePath = args[0];
            }
            else
            {
                Program.ErrorExit( $"Invalid number of args passed in. Expected 1, got {args.Length}\nBug glenn to add support for multi file parsing" );
            }

            SourceFile sourceFile = new SourceFile( filePath );

            string fileName = Path.GetFileName( filePath );
            Glossary glossary = new Glossary( fileName );

            while( sourceFile.HasData() )
            {
                glossary.AddEntry( sourceFile.ReadEntry() );
            }
            

            //Output to log for now, nextup protobuf or similar for export
            Console.WriteLine();
            foreach( KeyValuePair<string, string> kvp in glossary.Entries )
            {
                Console.WriteLine( $"StringID: {kvp.Key}\nLocString: {kvp.Value}\n" );
            }
            Console.WriteLine( $"Found {glossary.Entries.Count} entries in {glossary.Name}" );


            if( glossary.ErrorMessages.Count > 0 )
            {
                Console.WriteLine( $"{glossary.ErrorMessages.Count} issues were reported" );
                foreach( string str in glossary.ErrorMessages )
                {
                    Console.WriteLine( str );
                }
            }
        }

        private static void PrintInfo()
        {
            Console.WriteLine( "Pass in an xlsx dictionary to parse and export" );

            //todo extend, support flags / configs (ex, warnings/errors on dupes or empty cells)
            //Environment.Exit(0);
        }

        public static void ErrorExit( string message, int errorCode = 1 )   //TODO, move this someplace better
        {
            Console.WriteLine( "\n~~~~~~~~~~~~~~~~~~~~~~" );
            Console.WriteLine( $"Error:\n{message}" );
            Environment.Exit( errorCode );
        }
    }
}
