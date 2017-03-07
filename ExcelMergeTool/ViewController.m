//
//  ViewController.m
//  ExcelMergeTool
//
//  Created by 马权 on 05/03/2017.
//  Copyright © 2017 马权. All rights reserved.
//

#import "ViewController.h"
#import <LibXL/LibXL.h>
#import "Masonry/Masonry.h"

static NSString * const kChooseInputFile = @"Choose Input File";
static NSString * const kChooseOutputFile = @"Choose Output File";

@interface ViewController () <NSOpenSavePanelDelegate>

{
    __weak IBOutlet NSButton *_chooseInputBtn;
    __weak IBOutlet NSButton *_chooseOutputBtn;
    __strong NSText *_inputText;
    __strong NSText *_outputText;
    
    __strong NSOpenPanel *_openPanel;
    NSURL *_inputFileURL;
    NSURL *_outFileURL;
}

@end

@implementation ViewController

- (void)viewDidLoad {
    [super viewDidLoad];
    
    _inputText = [[NSText alloc] initWithFrame:CGRectZero];
    _inputText.editable = NO;
    [self.view addSubview:_inputText];
    [_inputText mas_makeConstraints:^(MASConstraintMaker *make) {
        make.top.equalTo(_chooseInputBtn);
        make.left.equalTo(_chooseInputBtn.mas_right).offset(20);
        make.right.equalTo(self.view).offset(-20);
        make.height.lessThanOrEqualTo(@20);
    }];
    
    _outputText = [[NSText alloc] initWithFrame:CGRectZero];
    _outputText.editable = NO;
    [self.view addSubview:_outputText];
    [_outputText mas_makeConstraints:^(MASConstraintMaker *make) {
        make.top.equalTo(_chooseOutputBtn);
        make.left.equalTo(_chooseOutputBtn.mas_right).offset(20);
        make.right.equalTo(self.view).offset(-20);
        make.height.lessThanOrEqualTo(@20);
    }];
}

- (IBAction)chooseInputAction:(id)sender {
    _openPanel = [NSOpenPanel openPanel];
    _openPanel.title = kChooseInputFile;
    _openPanel.allowsMultipleSelection = YES;
    _openPanel.canChooseFiles = YES;
    _openPanel.canChooseDirectories = YES;
    _openPanel.delegate = self;
    [_openPanel beginWithCompletionHandler:^(NSInteger result) {
        
    }];
}

- (IBAction)chooseOutputAction:(id)sender {
    _openPanel = [NSOpenPanel openPanel];
    _openPanel.title = kChooseOutputFile;
    _openPanel.allowsMultipleSelection = NO;
    _openPanel.canChooseFiles = YES;
    _openPanel.canChooseDirectories = YES;
    _openPanel.delegate = self;
    [_openPanel beginWithCompletionHandler:^(NSInteger result) {
        
    }];
}

- (BOOL)panel:(id)sender validateURL:(NSURL *)url error:(NSError **)outError
{
    NSOpenPanel *openPanel = nil;
    if ([sender isKindOfClass:[NSOpenPanel class]]) {
        openPanel = (NSOpenPanel *)sender;
    }
    if ([openPanel.title isEqualToString:kChooseInputFile]) {
        _inputFileURL = openPanel.URLs.firstObject;
        _inputText.string = _inputFileURL.path;
    }
    if ([openPanel.title isEqualToString:kChooseOutputFile]) {
        _outFileURL = openPanel.URLs.firstObject;
        _outputText.string = _outFileURL.path;
    }
    return YES;
}

- (IBAction)readExcel:(id)sender {
    BookHandle outputBookHandle = [[self class] bookHandleFilePath:_outFileURL.path];
    SheetHandle outputSheetHandle = xlBookGetSheet(outputBookHandle, 0);
    NSMutableArray<NSString *> *keysArray = [[self class] titleArrayWithSheetHandle:outputSheetHandle];

    BookHandle inputBookHandle = [[self class] bookHandleFilePath:_inputFileURL.path];
    NSMutableArray<NSDictionary *> *sheetsArray = @[].mutableCopy;
    int sheetCount = xlBookSheetCount(inputBookHandle);
    for (int sheetIndex = 0; sheetIndex < sheetCount; sheetIndex++) {
        SheetHandle sheetHandle = xlBookGetSheet(inputBookHandle, sheetIndex);
        int sheetRowCount = xlSheetLastRow(sheetHandle);
        NSMutableDictionary *sheetDic = @{}.mutableCopy;
        for (int row = 0; row < sheetRowCount; row++) {
            const char *cKey = xlSheetReadStr(sheetHandle, row, 0, NULL);
            if (cKey != NULL) {
                NSString *key = [NSString stringWithUTF8String:cKey];
                if ([[key substringFromIndex:key.length - 1] isEqualToString:@":"]) {
                    key = [key substringToIndex:key.length - 1];
                }
                if ([key isEqualToString:@"IMEI"]) {
                    NSLog(@"123");
                }
                if ([keysArray containsObject:key]) {
                    const char *cValue = xlSheetReadStr(sheetHandle, row, 1, NULL);
                    if (cValue != NULL) {
                        NSString *value = [NSString stringWithUTF8String:cValue];
                        [sheetDic setValue:value forKey:key];
                    }
                }
            }
        }
        if (sheetDic.allKeys.count > 0) {
            [sheetsArray addObject:sheetDic];
        }
    }
    
    __block int rowCount = xlSheetLastRowA(outputSheetHandle);
    [sheetsArray enumerateObjectsUsingBlock:^(NSDictionary * _Nonnull keyValueDic, NSUInteger idx, BOOL * _Nonnull stop) {
        [keysArray enumerateObjectsUsingBlock:^(NSString * _Nonnull key, NSUInteger col, BOOL * _Nonnull stop) {
            if ([[key substringFromIndex:key.length - 1] isEqualToString:@":"]) {
                key = [key substringToIndex:key.length - 1];
            }
            NSString *value = keyValueDic[key];
            xlSheetWriteStr(outputSheetHandle, rowCount, (int)col + 1, [value UTF8String], NULL);
        }];
        rowCount++;
    }];
    
    //  填充头部的列
    xlSheetInsertRow(outputSheetHandle, 1, 1);
    
    [keysArray enumerateObjectsUsingBlock:^(NSString * _Nonnull obj, NSUInteger idx, BOOL * _Nonnull stop) {
        xlSheetWriteStr(outputSheetHandle, 1, (int)idx + 1, [obj UTF8String], NULL);
    }];

    [[self class] saveBookHandle:outputBookHandle withName:@"new.xlsx"];
    
    xlBookRelease(outputBookHandle);
    xlBookRelease(inputBookHandle);
}

+ (NSMutableArray *)titleArrayWithSheetHandle:(SheetHandle)sheetHandle
{
    NSMutableArray *titleArray = @[].mutableCopy;
    int colCount = xlSheetLastColA(sheetHandle);
    for (int col = 1; col < colCount; col++) {
        const char * cTitle = xlSheetReadStr(sheetHandle, 0, col, NULL);
        if (cTitle != NULL) {
            NSString *title = [NSString stringWithUTF8String:cTitle];
            [titleArray addObject:title];
        }
    }
    return titleArray;
}

+ (void)saveBookHandle:(BookHandle)bookhandle
              withName:(NSString *)name
{
    xlBookSaveA(bookhandle, [[self filePathWithName:name] UTF8String]);
}

+ (BookHandle)bookHandleFilePath:(NSString *)filePath
{
    BookHandle handle = NULL;
    NSString *extension = [filePath pathExtension];
    if ([extension isEqualToString:@"xlsx"]) {
        handle = xlCreateXMLBook();
    }
    if ([extension isEqualToString:@"xls"]) {
        handle = xlCreateBook();
    }
    
    xlBookLoad(handle, [filePath UTF8String]);
    return handle;
}

+ (NSString *)filePathWithName:(NSString *)name
{
    NSString *documentPath = [NSSearchPathForDirectoriesInDomains(NSDocumentDirectory,NSUserDomainMask, YES) objectAtIndex:0];
    NSString *filePath = [documentPath stringByAppendingPathComponent:name];
    return filePath;
}

@end
