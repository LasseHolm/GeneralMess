unit Dpscad_TLB;

// ************************************************************************ //
// WARNING
// -------
// The types declared in this file were generated from data read from a
// Type Library. If this type library is explicitly or indirectly (via
// another type library referring to this type library) re-imported, or the
// 'Refresh' command of the Type Library Editor activated while editing the
// Type Library, the contents of this file will be regenerated and all
// manual modifications will be lost.
// ************************************************************************ //

// $Rev: 52393 $
// File generated on 05-04-2017 10:18:07 from Type Library described below.

// ************************************************************************  //
// Type Lib: C:\DelphiX\TRUNK\Main\scr\Dpscad (1)
// LIBID: {C5E8FDA4-C88D-11D1-AE4B-0020AF16D64A}
// LCID: 0
// Helpfile:
// HelpString: PC-SCHEMATIC Automation Library (v18)
// DepndLst:
//   (1) v2.0 stdole, (C:\Windows\SysWOW64\stdole2.tlb)
// SYS_KIND: SYS_WIN32
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers.
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
{$ALIGN 4}

interface

uses Winapi.Windows, System.Classes, System.Variants, System.Win.StdVCL, Vcl.Graphics, Vcl.OleServer, Winapi.ActiveX;


// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:
//   Type Libraries     : LIBID_xxxx
//   CoClasses          : CLASS_xxxx
//   DISPInterfaces     : DIID_xxxx
//   Non-DISP interfaces: IID_xxxx
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  DpscadMajorVersion = 2;
  DpscadMinorVersion = 7;

  LIBID_Dpscad: TGUID = '{C5E8FDA4-C88D-11D1-AE4B-0020AF16D64A}';

  DIID_IPCsFigureEvents: TGUID = '{281F4EB6-5696-4B49-ACEE-F3C0F1C4F6E3}';
  IID_IPCsFigure: TGUID = '{381FE3A1-B600-41D8-99BB-611897416A12}';
  IID_IPCsBase: TGUID = '{970CEB0E-ACAA-4E3F-B5B1-12833D73FF3B}';
  IID_IPCsLines: TGUID = '{DDE0BE45-1FBB-4729-AFB0-314BC3CEA689}';
  IID_IPCsLine: TGUID = '{13AECE11-B1EA-459C-BE19-7F484E12D3EE}';
  IID_IPCsPoints: TGUID = '{647664A1-8E29-4469-ACE8-26608477DF8C}';
  CLASS_PCsPoints: TGUID = '{A92CFB40-8A4B-44F2-AAB7-E4A023EE5636}';
  IID_IPCsStatusPoint: TGUID = '{28DF9ED4-686A-4317-8843-4EFAE1CD3C5F}';
  IID_IPCsJoinPoint: TGUID = '{3FB4759B-B7DE-4801-9375-6872D0B4A4A8}';
  IID_IPCsObject: TGUID = '{0896A10E-AF2C-450B-ACC0-32B4D0FD52C6}';
  IID_IPCsLayer: TGUID = '{60A0B67B-DCBD-4ADE-8E39-12761C092AAC}';
  IID_IPCsDrawing: TGUID = '{61BFFC36-9B1E-4BFC-92ED-EB532CD10268}';
  IID_IPCsPages: TGUID = '{DAC4E978-E186-4E2F-925F-E796EC3442D6}';
  IID_IPCsPage: TGUID = '{5771AA42-30B5-496E-AB85-FB3CCE6F69B4}';
  IID_IPCsSymbols: TGUID = '{FF44C293-724F-4E91-AA14-17825A792DEC}';
  IID_IPCsLibType: TGUID = '{EF5253F8-0F9C-47A5-8BA7-2D369B12A62D}';
  IID_IPCsLibTypes: TGUID = '{4C394DA5-F3F7-49AE-BFA0-84899BBB6D25}';
  IID_IPCsTexts: TGUID = '{81F82835-4133-46BA-8547-00BDB5D73C8F}';
  IID_IPCsSymbolTexts: TGUID = '{A6BD811D-303B-43D2-8942-EEDCC531EEBA}';
  IID_IPCsText: TGUID = '{17859DD2-080B-4256-8F37-CFD149D1E6B7}';
  IID_IPCsTextFont: TGUID = '{506D1B55-9E93-4FE7-9DC8-BDAEA92264B9}';
  IID_IPCsNameText: TGUID = '{EDE5A210-B27A-44AA-BBAF-1BE2702A372A}';
  IID_IPCsSymbol: TGUID = '{9F12909F-4EE6-40C7-89C7-5B110A48355A}';
  IID_IPCsConnections: TGUID = '{9CC2A147-2E7D-4613-B505-10B5AAD7AD2D}';
  IID_IPCsConnection: TGUID = '{C5A27EFD-744C-4DCD-B4B5-1E4AC44023C7}';
  IID_IPCsConnectionTexts: TGUID = '{9C483B8F-5BE6-4675-BDDA-8984012826B8}';
  IID_IPCsDatafields: TGUID = '{F4E0CF05-3AA2-4D46-B9AD-C199940CE630}';
  IID_IPCsDatafield: TGUID = '{90B2DFB9-8544-4BFE-9ECD-3E23B69ECF7D}';
  IID_IPCsDataLink: TGUID = '{A07DB9EC-882F-4041-8B69-A16E8AD3BA8D}';
  IID_IPCsPageSetup: TGUID = '{8462D121-9F59-4C75-B741-499FEF19488C}';
  IID_IPCsPageReference: TGUID = '{A845D055-E632-48E9-B3C6-BCA6C2772518}';
  IID_IPCsBasePageRef: TGUID = '{7C95FC24-0E20-4471-B5A7-DE41C4FC9772}';
  IID_IPCsExternalObjects: TGUID = '{AE4237D8-E4BE-427B-AC4E-14D3EB492067}';
  IID_IPCsLayers: TGUID = '{FEAF952B-B3F0-4671-AF8B-2026453A177E}';
  IID_IPCsApplication: TGUID = '{226A846F-E175-4CE5-8C6F-B57EC56EEAB6}';
  IID_IPCsDocuments: TGUID = '{5731AB31-7CA2-45F2-B84C-108E708E7FD0}';
  IID_IPCsDocument: TGUID = '{D7D52290-DB58-45FA-85BB-25FB073FA5AA}';
  IID_IPCsApplicationPreferences: TGUID = '{0042104A-A21F-4DFB-A60A-A078BA62E955}';
  IID_IPCsApplicationPreferences2: TGUID = '{106E13A7-E58E-4A3A-A257-F88C3029F1FD}';
  IID_IPCsSymbolMenu: TGUID = '{5ED0A430-81EC-4529-9D9B-A35EF0A3177E}';
  IID_IPCsComponentDatabase: TGUID = '{C8B26975-AC5E-473F-B922-DA08C50C383E}';
  IID_IPCsLibTypeConnections: TGUID = '{0573CF73-2CD7-4B66-BABA-4421EC53A302}';
  IID_IPCsArcs: TGUID = '{64AF72DA-7A46-4033-BFE9-034BEEC408FF}';
  IID_IPCsArc: TGUID = '{BC2F1F85-76B3-408D-860D-215F963B9A40}';
  IID_IApp: TGUID = '{C5E8FDA5-C88D-11D1-AE4B-0020AF16D64A}';
  IID_IPCsApplicationObsolete: TGUID = '{6CF6C0C6-D8FB-11D5-9030-0020AF16D6BF}';
  DIID_IPCsApplicationEvents: TGUID = '{4576D15E-0EE2-4B1B-8AC3-2E820FEE1222}';
  IID_IPCsApplication2: TGUID = '{F6256177-1C51-460F-84C7-AB1905D508E5}';
  CLASS_PCsDocuments: TGUID = '{B344FD29-EC43-423B-893B-EA70C9984730}';
  DIID_IPCsDocumentEvents: TGUID = '{0061C56D-4338-4314-99DD-FB039E7BB236}';
  IID_IPCsDocument2: TGUID = '{43F030B6-1F4C-47CE-9A07-98BC1ECFAE2F}';
  DIID_IPCsApplicationPreferencesEvents: TGUID = '{836113FE-0903-4C9D-B517-33BFBD5B7191}';
  CLASS_PCsApplicationPreferences: TGUID = '{25377BED-6499-41EC-B327-F8DE10701041}';
  DIID_IPCsStatusPointEvents: TGUID = '{2DA70029-2E21-4962-A583-A4034CD9D819}';
  CLASS_PCsStatusPoint: TGUID = '{7A71BA8F-4FE9-4604-B905-8F7FEF828E1B}';
  IID_IPCsLibType2: TGUID = '{7203263B-5703-4251-B3A4-827118605237}';
  DIID_IPCsDrawingEvents: TGUID = '{E069BDF3-AE9A-43A6-902F-6D8CF814E1D4}';
  IID_IPCsDrawing2: TGUID = '{9CD75DDE-BDF9-446F-BB19-3BEDF942A0A1}';
  DIID_IPCsLibTypesEvents: TGUID = '{DB192BAD-CAB3-45D6-AC9D-B92A33CADB48}';
  CLASS_PCsLibTypes: TGUID = '{58552579-068D-49DF-A9C9-1A05760C59E0}';
  DIID_IPCsPagesEvents: TGUID = '{DD56960D-C565-48AE-812B-3B03D1D49E8B}';
  IID_IPCsPages2: TGUID = '{C4DE62DA-2FAA-4395-BA73-6B5C6BD16232}';
  IID_IPCsSymbol2: TGUID = '{669A39E0-9570-4494-BBB8-2AF3C5BF9249}';
  IID_IPCsSymbolsEnumerable: TGUID = '{F3B79EBD-7D7C-4550-9CBD-BB5179471DFC}';
  CLASS_PCsLines: TGUID = '{D8DCBDBB-D727-48D2-B0C5-E75AF6572BCE}';
  IID_IPCsLine2: TGUID = '{7653E66C-60EC-467A-90B3-2376CBEFA0FC}';
  DIID_IPCsPageEvents: TGUID = '{5D80BF9D-FFAC-4B17-B749-2E670C4CAD64}';
  IID_IPCsPage2: TGUID = '{D0465CE5-FBE4-4CDC-9457-991F4C1FBE45}';
  IID_IPCsArc2: TGUID = '{02F30A6A-EC65-4B42-9189-31946D8D6CB0}';
  CLASS_PCsArcs: TGUID = '{7253EEE5-24CC-42C5-A1C9-95EA7D104D7F}';
  CLASS_PCsPageSetup: TGUID = '{75CBB0D1-5C5A-4A77-A077-31A34963AC7E}';
  CLASS_PCsPageReference: TGUID = '{5890B5FC-45F1-4F0D-9666-66AE3378CC18}';
  CLASS_PCsBasePageRef: TGUID = '{37DA9814-4B50-48CC-A643-1293C13C17D7}';
  CLASS_PCsTexts: TGUID = '{65DD5A53-DB57-4F3F-8FCC-F572E94EE50C}';
  CLASS_PCsSymbolTexts: TGUID = '{D01F5112-44EC-47BE-AB11-BA6A84927353}';
  CLASS_PCsSymbolMenu: TGUID = '{E08F1A6D-A990-44B7-9C82-CEB215000E07}';
  CLASS_PCsLibTypeConnections: TGUID = '{0A12B13A-03BE-4841-B3A5-820218834FD4}';
  CLASS_PCsConnections: TGUID = '{5C26B6D4-D009-4C64-88F4-6532F3E5258B}';
  CLASS_PCsConnectionTexts: TGUID = '{B17309C1-144A-45AC-B1DA-7D5507DD6073}';
  IID_IPCsStatusObject: TGUID = '{E78749F9-B862-46B4-B3D9-FDC143E67FDF}';
  CLASS_PCsTextFont: TGUID = '{76D9939F-8988-4264-BD13-3999F9C37B26}';
  IID_IPCsConnection2: TGUID = '{94CF8C33-E3F3-49CE-B9D9-D695918670EF}';
  CLASS_PCsJoinPoint: TGUID = '{22C27040-41ED-4182-A345-53A01B9CA6AE}';
  CLASS_PCsLayers: TGUID = '{A6497A2B-2671-48D8-A301-4A6B6D8B91A0}';
  CLASS_PCsLayer: TGUID = '{28AA17DF-01E7-47E1-A887-35F41B18AE19}';
  CLASS_PCsObject: TGUID = '{913E4614-EA80-478E-A6C0-CC857BDC5E44}';
  IID_IPCsComponentDatabase2: TGUID = '{FE0869BA-D134-4ED5-80ED-E57EFA8A4CF6}';
  CLASS_App: TGUID = '{C5E8FDA6-C88D-11D1-AE4B-0020AF16D64A}';
  IID_IPCsGlobal: TGUID = '{C6BF0C78-7D73-4D6D-B833-743580EE5B5D}';
  IID_IPCsGlobal2: TGUID = '{200950A3-3BE4-415A-B5C9-10D3D5A340C8}';
  IID_IPCsExternalObjects2: TGUID = '{E42E13EE-7A40-4C50-9FDE-F41360A4A199}';
  IID_IPCsDatafields2: TGUID = '{4268FFB3-BE93-47E6-A20D-E898BFEE0467}';
  CLASS_PCsDatafield: TGUID = '{7B49979D-FE87-4E04-ABCE-EE323F33B653}';
  CLASS_PCsDataLink: TGUID = '{E1479499-9BA1-41E9-A655-6C2BA3AD1FEB}';
  IID_IPCsNameText2: TGUID = '{32E89920-84E7-474A-BC7E-E53AE9AAC79F}';
  IID_IPCsReferenceDesignations: TGUID = '{42F4159D-C012-4641-9D20-A740E69BF862}';
  IID_IPCsReferenceDesignations2: TGUID = '{3FB35D3D-29E2-48D9-B6DA-687710CC743C}';
  IID_IPCsReferenceDesignation: TGUID = '{03DA2962-6DA8-47A0-85BA-279314C3E075}';
  CLASS_PCsReferenceDesignation: TGUID = '{98821517-773D-41D4-828C-FDFDB1A33D66}';
  IID_IPCsSymbolFolderAliases: TGUID = '{F70942D5-008E-4E4D-B415-E305833B7EF4}';
  CLASS_PCsSymbolFolderAliases: TGUID = '{0A00A357-681C-4170-9241-228FDFC52554}';
  IID_IPCsApplication3: TGUID = '{DEE2514B-736F-4A32-AD5C-35E6C7373407}';
  IID_IPCsDialog: TGUID = '{C7485B3E-E4D1-49C0-AD43-BD4E390FC792}';
  DIID_IPCsDialogEvents: TGUID = '{1370ECF3-341D-43F9-9B70-0C111D611035}';
  CLASS_PCsDialog: TGUID = '{12A58486-4A3D-4A25-BD4E-A3B2C016229E}';
  IID_IPCsFileDialog: TGUID = '{5C3A5FE5-91C7-44ED-BE25-EEE64F549345}';
  DIID_IPCsFileDialogEvents: TGUID = '{6525E6D9-BCD5-4E09-AE66-E5933DA8227A}';
  CLASS_PCsFileDialog: TGUID = '{E8830229-D4F0-4D2C-92F6-64E9CBF39BCF}';
  IID_IPCsLineTexts: TGUID = '{0423A4FD-AD69-406C-B337-2999E0FA7AE0}';
  CLASS_PCsLineTexts: TGUID = '{CDDEC1D5-8B08-4F05-ABB0-E47A164F8771}';
  IID_IPCsLists: TGUID = '{A7AEB7AD-3A64-43FF-9836-C9663D974EF2}';
  IID_IPCsLists2: TGUID = '{5321BF67-C846-444E-ABB6-36CF6CA8C0F3}';
  IID_IPCsList: TGUID = '{32055A0A-D65C-4A85-A964-E00403423559}';
  IID_IPCsList2: TGUID = '{F38CB7E9-64C0-47D6-A484-40E4378F8623}';
  IID_IPCsExplorerWindow: TGUID = '{E3997986-4A75-46F5-9B19-A9C2FE24385B}';
  DIID_IPCsExplorerWindowEvents: TGUID = '{816A2A2D-1F6A-47BD-A36C-FB86AF0EC7AB}';
  CLASS_PCsExplorerWindow: TGUID = '{3888DC6A-7BA5-45D5-A1C8-ABC56E31FDA6}';
  IID_IPCsSubDrawing: TGUID = '{0BB8A6D4-2E7D-48D6-9953-CE79FB2AAC33}';
  CLASS_PCsSubDrawing: TGUID = '{F0BBA5E0-D994-46B0-95F9-E2A8736C0FB7}';
  IID_IPCsListConnData: TGUID = '{AA0AC284-ACED-46F6-B287-0FCAA2B85BBB}';
  CLASS_PCsListConnData: TGUID = '{613CFD4D-82CC-4AED-BC7C-BE25405CA886}';
  IID_IPCsSelectStatus: TGUID = '{FA88F8B0-26F1-441C-BBA0-901106CCB211}';
  DIID_IPCsSelectStatusEvents: TGUID = '{453EB166-95DD-432B-A435-1A68BF509B8E}';
  CLASS_PCsSelectStatus: TGUID = '{FFE590C4-D8BB-4DD3-B697-7995FAF25C2C}';
  IID_IPCsSelectionArea: TGUID = '{755A38CB-812B-41CF-9BA6-696F96088873}';
  CLASS_PCsSelectionArea: TGUID = '{1624B860-7AA0-404F-B34D-5383DCD5CCBB}';
  IID_IPCsJoins: TGUID = '{13FA989D-485B-4BE2-9062-148F69DA8D33}';
  CLASS_PCsJoins: TGUID = '{4EB5D98C-8845-439D-B758-592EB020FD56}';
  IID_IPCsJoin: TGUID = '{5B281FCC-D09C-4C0C-935C-EFE272D57E0F}';
  CLASS_PCsJoin: TGUID = '{E138DAF2-889B-4274-ADE0-AD0EDF98FEC8}';
  IID_IPCsDrawingPreferences: TGUID = '{66BFC81B-1CCA-4E53-8C15-51A80752534B}';
  CLASS_PCsDrawingPreferences: TGUID = '{D23AF8B1-F3FB-4FDC-ADFA-CAC802944E9F}';
  IID_IPCsReference: TGUID = '{3077DBD0-D206-4478-A21F-E709AC5D1EDE}';
  CLASS_PCsReference: TGUID = '{455F44C1-F3CF-4930-A631-F714432EFACB}';
  IID_IPCsReferenceFrame: TGUID = '{7FED85F3-51A8-4B62-80BA-6F615B6206FF}';
  CLASS_PCsReferenceFrame: TGUID = '{1166E0E6-B5A8-4348-B385-CEDEC89DB498}';
  CLASS_PCsDocument: TGUID = '{62D5AF39-C3C0-4B10-B199-39C0875ABEEC}';
  IID_IPCsSymbol3: TGUID = '{68FA7365-D511-43B4-B220-B01EA114EA8A}';
  CLASS_PCsFigure: TGUID = '{2F81CBED-9B2E-4EC4-9150-2E5630556852}';
  CLASS_PCsList: TGUID = '{042E8E45-87C6-4064-B540-3EC62E36118B}';
  IID_IPCsLine3: TGUID = '{41776201-3CAF-4CB0-A10F-D9DE66846CBF}';
  CLASS_PCsLibType: TGUID = '{12FF19A9-DA97-4AB5-95CF-73F63CBFAFDC}';
  IID_IPCsPage3: TGUID = '{2A6D4CCB-9F5E-4B79-A285-BF1A375320D6}';
  IID_IPCsDrawing3: TGUID = '{24AB84D8-EB62-4160-8238-C5FD95598D4F}';
  IID_IPCsGlobal3: TGUID = '{212D442E-0E51-45DF-809F-DAB5397B80D0}';
  IID_IPCsDrawingPreferences2: TGUID = '{BA0AFD46-55FF-48BA-87F2-5FC856AE81C7}';
  IID_IPCsNameText3: TGUID = '{9305B2C0-5267-4A89-8AA5-06A5842C07C1}';
  IID_IPCsSubDrawing2: TGUID = '{6B8CCDD1-805E-43A9-8B54-0444DA25F734}';
  IID_IPCsTemplateData: TGUID = '{09646609-A0D9-49A5-95B3-483E77F7E479}';
  CLASS_PCsTemplateData: TGUID = '{30B083B7-7136-4F34-A7AD-ADBEC6BEA06F}';
  CLASS_PCsPages: TGUID = '{976F1EBD-4D61-458C-B532-6796750F5D83}';
  IID_IPCsExternalObject: TGUID = '{E51DE5D6-F827-476C-A50A-08BDA31F3257}';
  DIID_IPCsExternalObjectEvents: TGUID = '{D703C1D4-B497-4CCE-8EF7-24BD10DDBAAF}';
  IID_IPCsExternalObject2: TGUID = '{71A0B111-181F-4F28-9704-DFFF161FA5D8}';
  IID_IPCsPicture: TGUID = '{8985BD82-9EDB-4674-84BB-32A2D457299D}';
  DIID_IPCsPictureEvents: TGUID = '{6D87C31D-591E-4332-87B7-C4E49011266D}';
  CLASS_PCsPicture: TGUID = '{F37D8C1D-E994-45AB-B597-CC1BCB68823F}';
  IID_IPCsDatabaseSetup: TGUID = '{A843D727-CABB-4842-90B6-FA063DEDD883}';
  CLASS_PCsDatabaseSetup: TGUID = '{79B5512C-BF5E-4EC5-9BF3-7D2B41A23A9D}';
  IID_IPCsLine4: TGUID = '{2865F44C-BA03-4D85-A263-8FF3C7B6D841}';
  CLASS_Application: TGUID = '{6CF6C0CA-D8FB-11D5-9030-0020AF16D6BF}';
  IID_IPCSLine5: TGUID = '{DD427251-8AB8-4A46-A137-9BBBE3316984}';
  IID_IPCsSymbol4: TGUID = '{5E752712-E638-42D9-86EE-6002653969A7}';
  CLASS_PCsDrawing: TGUID = '{B955AA54-47D8-496D-8319-48DCEC92E64C}';
  CLASS_PCsComponentDatabase: TGUID = '{8645DC89-9501-4EBF-A7CD-49DD65164AC6}';
  CLASS_PCsLists: TGUID = '{92E2E8B9-0C70-4B7E-BB80-8D5EE926646B}';
  IID_IPCsPage4: TGUID = '{8CFB2E38-7691-4120-93AF-F0D48EAB9EFE}';
  IID_IPCsText2: TGUID = '{36BAF7D3-3BB8-4F52-9E85-6CD291ED2CA9}';
  CLASS_PCsStatusObject: TGUID = '{B399D42C-D6E0-4D4F-8986-37DA309387DC}';
  CLASS_PCsConnection: TGUID = '{A289BCF1-E902-44B2-96B3-BA3E24D5B2B1}';
  CLASS_PCsExternalObjects: TGUID = '{781F4867-F664-48AD-B2DA-CA74A60DDEDF}';
  CLASS_PCsExternalObject: TGUID = '{12D525B7-1AD3-4F2F-A4AA-478D634AA5AC}';
  IID_IPCsNameText4: TGUID = '{59F989D2-211C-48DF-9B25-D445CD7C1C01}';
  IID_IPCsComponent: TGUID = '{8DC3B9F5-5E29-48CE-8DBE-C700B1280769}';
  DIID_IPCsComponentEvents: TGUID = '{6DB9B522-43A8-499F-A47E-6CB900911C56}';
  IID_IPCsComponent2: TGUID = '{9D293E30-A1E2-47AB-9E73-DD99425AE035}';
  IID_IPCsComponentName: TGUID = '{7CD5EEA6-542C-48AF-894F-7D743883E37C}';
  CLASS_PCsComponentName: TGUID = '{0833CA02-5917-4708-BCEC-501C39792463}';
  CLASS_PCsSymbol: TGUID = '{54868618-5F62-441F-92C7-6B09B37DB24E}';
  IID_IPCsObjectRefID: TGUID = '{92E3FB02-7A93-4CEB-BB4B-E74E2DD7663A}';
  CLASS_PCsObjectRefID: TGUID = '{428B7F62-4CF8-45C6-A767-7584141D9B55}';
  IID_IPCsReferenceDesignations3: TGUID = '{8500668B-9672-416D-93C1-E55856AB7CAA}';
  IID_IPCsReferenceDesignationList: TGUID = '{89CE5B73-CFDF-40BD-BF28-9B00E5B56E48}';
  CLASS_PCsReferenceDesignationList: TGUID = '{71B62312-721E-416D-B4BA-D35F2B5D317A}';
  CLASS_PCsText: TGUID = '{635C5054-760C-4466-958A-E6E8BC306444}';
  IID_IPCsLine6: TGUID = '{99E2962E-0B32-44B3-B249-BFBAF6FC7822}';
  CLASS_PCsArc: TGUID = '{770453A0-8CC8-4C42-B118-C6E899AD3154}';
  CLASS_PCsSymbols: TGUID = '{5715688B-516F-4434-AD35-BF3F11F063CA}';
  IID_IPCsSymbolsEnumerator: TGUID = '{A52328DD-0BBB-4923-AA9E-B7258F6FCB75}';
  IID_IPCsComponentSymbols: TGUID = '{C1E7B59A-00DF-45CF-93D2-A665A0F4A5A7}';
  CLASS_PCsComponentSymbols: TGUID = '{712AAA0F-BD3E-4E6D-92DC-1B9DB8D1EE8D}';
  CLASS_PCsSymbolsEnumerator: TGUID = '{B0270391-10B4-463E-9C0F-492F55CF9CC5}';
  IID_IPCsPage5: TGUID = '{2445EFE9-48A5-4CB2-A661-0C19EB145822}';
  CLASS_PCsLine: TGUID = '{39F2E948-4484-4725-8E19-BC14201CEE17}';
  IID_IPCsLineDatafields: TGUID = '{4FD608C3-5B5C-42E1-B5E9-4021660091B1}';
  CLASS_PCsDatafields: TGUID = '{FE88AD28-EFBF-4A3F-BC37-11C9D67A79FD}';
  CLASS_PCsLineDatafields: TGUID = '{1968B1DC-3692-4F2C-AC08-8A1560F9747B}';
  CLASS_PCsNameText: TGUID = '{83B13CFA-433E-41A0-BB2B-A7D0A47CAC4E}';
  IID_IPCsRefIDFormat: TGUID = '{A40E9EA5-B27B-4FC8-9470-C036E0C1FDB0}';
  CLASS_PCsRefIDFormat: TGUID = '{8DC0886E-2B88-4081-BB8A-AF5920B431E2}';
  CLASS_PCsGlobal: TGUID = '{EB175FFF-2D77-40C2-AE35-F305475526C0}';
  CLASS_PCsReferenceDesignations: TGUID = '{75D62F60-4A06-4051-8ED3-438376B73E45}';
  CLASS_PCsPage: TGUID = '{FBCA35D8-E6C5-43E9-8396-0C75F74EC77E}';
  CLASS_PCsComponent: TGUID = '{702F2DAA-8164-4849-A727-DB07D69A2319}';
  IID_IPCsCompName: TGUID = '{0C073709-BCC6-4134-A670-6B2F61879C3E}';
  CLASS_PCsCompName: TGUID = '{9C887473-CE2D-423E-934C-46C0533C76F8}';

// *********************************************************************//
// Declaration of Enumerations defined in Type Library
// *********************************************************************//
// Constants for enum psPageRectFlag
type
  psPageRectFlag = TOleEnum;
const
  pr_Header = $00000002;
  pr_Texts = $00000004;
  pr_Origo = $00000010;

// Constants for enum psIOStatusExtension
type
  psIOStatusExtension = TOleEnum;
const
  io_None = $00000000;
  io_All = $00000001;
  io_Term = $00000002;
  io_PLC = $00000003;

// Constants for enum psIOStatusMainType
type
  psIOStatusMainType = TOleEnum;
const
  mt_None = $00000000;
  mt_Output = $00000001;
  mt_Input = $00000002;

// Constants for enum PsLineStyle
type
  PsLineStyle = TOleEnum;
const
  ls_Solid = $00000000;
  ls_Dash = $00000001;
  ls_Dot = $00000002;
  ls_DashDot = $00000003;
  ls_DashDotDot = $00000004;
  ls_DashDotDotDot = $00000005;
  ls_DashDashDot = $00000006;
  ls_DashDashDotDot = $00000007;
  ls_DashDashDotDotDot = $00000008;
  ls_FinDot = $00000009;
  ls_DualColor = $0000000A;
  ls_Free3 = $0000000B;
  ls_Free4 = $0000000C;
  ls_Free5 = $0000000D;
  ls_SolidBox = $0000000E;
  ls_EmptyBox = $0000000F;
  ls_FDiagonal05Box = $00000010;
  ls_FDiagonal10Box = $00000011;
  ls_BDiagonal05Box = $00000012;
  ls_BDiagonal10Box = $00000013;
  ls_DiagCross05Box = $00000014;
  ls_VerticalBox = $00000015;
  ls_HorizontalBox = $00000016;
  ls_CrossBox = $00000017;
  ls_ZHashedBox = $00000018;
  ls_RoundBox = $00000019;
  ls_Free6 = $0000001A;
  ls_Free7 = $0000001B;
  ls_Free8 = $0000001C;
  ls_Free9 = $0000001D;
  ls_Free10 = $0000001E;
  ls_SolidFDiag50 = $0000001F;
  ls_SolidCross50 = $00000020;
  ls_Free11 = $00000021;
  ls_Free12 = $00000022;
  ls_Free13 = $00000023;
  ls_Free14 = $00000024;
  ls_Free15 = $00000025;
  ls_Free16 = $00000026;
  ls_DashDotChannel = $00000027;
  ls_DashDotChannel25 = $00000028;
  ls_DashDotChannel50 = $00000029;
  ls_Free17 = $0000002A;
  ls_ZFigur = $0000002B;
  ls_BDiagonal05 = $0000002C;
  ls_BDiagonal10 = $0000002D;
  ls_VertHoriBox = $0000002E;
  ls_DashedBox = $0000002F;

// Constants for enum PsPenColor
type
  PsPenColor = TOleEnum;
const
  Co_Black = $00000000;
  Co_Red = $00000001;
  Co_Green = $00000002;
  Co_Yellow = $00000003;
  Co_Blue = $00000004;
  Co_Magenta = $00000005;
  Co_Cyan = $00000006;
  Co_White = $00000007;
  Co_LRed = $00000008;
  Co_LGreen = $00000009;
  Co_Brown = $0000000A;
  Co_LBlue = $0000000B;
  Co_LMagenta = $0000000C;
  Co_LCyan = $0000000D;
  Co_NoPrint = $0000000E;
  Co_Dim = $0000000F;
  Co_HelpColor = $00000010;
  Co_WinColor = $00000011;
  co_BlackColor = $00000012;
  co_HighLight = $00000013;

// Constants for enum PsPCsObjectType
type
  PsPCsObjectType = TOleEnum;
const
  ot_Symbol = $00000000;
  ot_Signal = $00000001;
  ot_MultiSignal = $00000002;
  ot_WireNumber = $00000003;
  ot_TestPoint = $00000004;
  ot_Cable = $00000005;
  ot_Connection = $00000006;
  ot_Line = $00000007;
  ot_Page = $00000008;
  ot_DrawSymb = $00000009;
  ot_Text = $0000000A;
  ot_RefFrame = $0000000B;
  ot_Insertion = $0000000C;
  ot_Measure = $0000000D;
  ot_LineConnect = $0000000E;
  ot_Branch = $0000000F;

// Constants for enum psSymbolType
type
  psSymbolType = TOleEnum;
const
  st_Normal = $00000000;
  st_Head = $00000001;
  st_Ignore = $00000002;
  st_Relay = $00000003;
  st_Open = $00000004;
  st_Close = $00000005;
  st_Switch = $00000006;
  st_MasterRef = $00000007;
  st_WithRef = $00000008;
  st_Reference = $00000009;
  st_Signal = $0000000A;
  st_Terminal = $0000000B;
  st_Plc = $0000000C;
  st_Data = $0000000D;
  st_NoneElectric = $0000000E;
  st_Support = $0000000F;
  st_Power = $00000010;
  st_Cable = $00000011;
  st_WireNumber = $00000012;
  st_TestPoint = $00000013;
  st_MultiSignal = $00000014;
  st_DynCable = $00000015;
  st_DynTerm = $00000016;
  st_PowerSource = $00000017;
  st_PowerDistrb = $00000018;
  st_PowerLoad = $00000019;
  st_PowerCable = $0000001A;
  st_CableDuct = $0000001B;
  st_Insertion = $0000001C;
  st_PlcRef = $0000001D;
  st_Measure = $0000001E;
  st_LineConnect = $0000001F;
  st_Branch = $00000020;
  st_BusBar = $00000021;
  st_PowerFuse = $00000022;

// Constants for enum PsTextType
type
  PsTextType = TOleEnum;
const
  tt_None = $00000000;
  tt_FreeText = $00000001;
  tt_DataField = $00000002;
  tt_SymbolName = $00000003;
  tt_SymbolType = $00000004;
  tt_SymbolArticle = $00000005;
  tt_SymbolFunction = $00000006;
  tt_PinName = $00000007;
  tt_PinFunction = $00000008;
  tt_PinLabel = $00000009;
  tt_PinDescription = $0000000A;
  tt_SignalRef = $0000000B;
  tt_SingleRef = $0000000C;
  tt_MultiRef = $0000000D;
  tt_XPageRef = $0000000E;
  tt_YPageRef = $0000000F;

// Constants for enum PsSymbolRectType
type
  PsSymbolRectType = TOleEnum;
const
  rt_Raw = $00000000;
  rt_WithTexts = $00000020;
  rt_WithTextsAlsoBlank = $00000060;
  rt_Relative = $00000080;
  rt_Reserved1 = $00000002;
  rt_BlankTexts = $00000040;
  rt_WithReference = $00000100;

// Constants for enum psOrientation
type
  psOrientation = TOleEnum;
const
  orLandscape = $00000000;
  orPortrait = $00000001;

// Constants for enum psPageType
type
  psPageType = TOleEnum;
const
  pt_Schematic = $00000000;
  pt_GroundPlane = $00000001;
  pt_ISOmetric = $00000002;
  pt_SemiISO = $00000003;
  pt_Ignore = $00000004;

// Constants for enum psPageFunction
type
  psPageFunction = TOleEnum;
const
  pf_Normal = $00000000;
  pf_Ignore = $00000001;
  pf_Unit = $00000002;
  pf_Index = $00000003;
  pf_Part = $00000004;
  pf_Power = $00000005;
  pf_Comp = $00000006;
  pf_Term = $00000007;
  pf_Cable = $00000008;
  pf_PLC = $00000009;
  pf_Net = $0000000A;
  pf_TermPlan = $0000000B;
  pf_CablePlan = $0000000C;
  pf_Divider = $0000000D;
  pf_ConnPlan = $0000000E;
  pf_SymbolDoc = $0000000F;

// Constants for enum psDirection
type
  psDirection = TOleEnum;
const
  dirNone = $00000000;
  dirHorizontal = $00000001;
  dirVertical = $00000002;

// Constants for enum PsFolderType
type
  PsFolderType = TOleEnum;
const
  ft_System = $FFFFFFFF;
  ft_Home = $FFFFFFFE;
  ft_Project = $00000000;
  ft_Symbol = $00000001;
  ft_Units = $00000002;
  ft_Database = $00000003;
  ft_Standard = $00000004;
  ft_List = $00000005;
  ft_Templates = $00000006;
  ft_Admin = $00000007;
  ft_PDF = $00000008;
  ft_Scripts = $00000009;
  ft_DataDefinitions = $0000000A;
  ft_FormatFiles = $0000000B;

// Constants for enum PsTextOrigin
type
  PsTextOrigin = TOleEnum;
const
  to_LB = $00000000;
  to_CB = $00000001;
  to_RB = $00000002;
  to_LC = $00000003;
  to_CC = $00000004;
  to_RC = $00000005;
  to_LT = $00000006;
  to_CT = $00000007;
  to_RT = $00000008;

// Constants for enum psTextRectFlags
type
  psTextRectFlags = TOleEnum;
const
  rf_BlankTexts = $00000040;
  rf_AllRect = $00000001;

// Constants for enum psReferenceDesignationAspect
type
  psReferenceDesignationAspect = TOleEnum;
const
  asFunction = $00000000;
  asLocation = $00000001;
  asProduct = $00000002;
  asFacility = $00000003;

// Constants for enum psSymbolRefDesOption
type
  psSymbolRefDesOption = TOleEnum;
const
  roShowFull = $00000002;
  roLineBreak = $00000004;
  roLineBreakInherited = $00000008;

// Constants for enum psDialogKind
type
  psDialogKind = TOleEnum;
const
  dkDialogFileOpen = $00000000;

// Constants for enum psListType
type
  psListType = TOleEnum;
const
  ltNone = $00000000;
  ltIndex = $00000001;
  ltPart = $00000002;
  ltComp = $00000003;
  ltTerm = $00000004;
  ltCable = $00000005;
  ltPLC = $00000006;
  ltNet = $00000007;
  ltConns = $00000008;

// Constants for enum psPlaceSubDrawingOption
type
  psPlaceSubDrawingOption = TOleEnum;
const
  oo_KeepPageNumbers = $00000001;
  oo_AutoName = $00000002;
  oo_AutoPlace = $00000004;

// Constants for enum psChangeDrawingOrder
type
  psChangeDrawingOrder = TOleEnum;
const
  doBringToFront = $00000000;
  doSendToBack = $00000001;
  doSendForward = $00000002;
  doSendBackward = $00000003;

// Constants for enum psGotoObjectOption
type
  psGotoObjectOption = TOleEnum;
const
  goKeepSelection = $00000002;
  goSelect = $00000004;

// Constants for enum psSetupDatabaseFieldTypes
type
  psSetupDatabaseFieldTypes = TOleEnum;
const
  dfArt = $00000000;
  dfAltArt = $00000001;
  dfType = $00000002;
  dfFunction = $00000003;
  dfSymbol = $00000004;
  dfSymbolRef = $00000005;
  dfPinName = $00000006;
  dfMecSymbol = $00000007;
  dfAdditionals = $00000008;
  dfPrice1 = $00000009;
  dfPrice2 = $0000000A;
  dfDiscount1 = $0000000B;
  dfDiscount2 = $0000000C;
  dfSearch = $0000000D;
  dfSearch2 = $0000000E;
  dfReadonly = $0000000F;
  dfMark = $00000010;
  dfPowerSymbol = $00000011;
  dfPowerPins = $00000012;
  dfManifacturer = $00000013;
  dfApproved = $00000014;
  dfObsolete = $00000015;
  dfDescription = $00000016;
  dfPackUnit = $00000017;
  dfSLDSymbol = $00000018;
  dfSubDrawing = $00000019;
  dfPBCurrent = $0000001A;
  dfPBPhases = $0000001B;
  dfPBModules = $0000001C;
  dfPBCompFunc = $0000001D;

// Constants for enum psHighlighting
type
  psHighlighting = TOleEnum;
const
  hl_None = $00000000;
  hl_Highlighting1 = $00000001;

// Constants for enum psSelectOptions
type
  psSelectOptions = TOleEnum;
const
  so_Replace = $00000000;
  so_Add = $00000001;

// Constants for enum psIterateDrawingObjectsOptions
type
  psIterateDrawingObjectsOptions = TOleEnum;
const
  ho_SkipSymbols = $00000001;
  ho_SkipConnections = $00000002;
  ho_SkipTexts = $00000004;
  ho_SkipLines = $00000008;

// Constants for enum psIsSameComponentOptions
type
  psIsSameComponentOptions = TOleEnum;
const
  sp_AssumeEqualSymbolNames = $00000001;
  sp_IgnoreArticle = $00000002;

// Constants for enum psMetaconnectivityOptions
type
  psMetaconnectivityOptions = TOleEnum;
const
  mo_ThroughTerminals = $00000001;
  mo_ThroughBusBars = $00000002;

type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary
// *********************************************************************//
  IPCsFigureEvents = dispinterface;
  IPCsFigure = interface;
  IPCsFigureDisp = dispinterface;
  IPCsBase = interface;
  IPCsBaseDisp = dispinterface;
  IPCsLines = interface;
  IPCsLinesDisp = dispinterface;
  IPCsLine = interface;
  IPCsLineDisp = dispinterface;
  IPCsPoints = interface;
  IPCsPointsDisp = dispinterface;
  IPCsStatusPoint = interface;
  IPCsStatusPointDisp = dispinterface;
  IPCsJoinPoint = interface;
  IPCsJoinPointDisp = dispinterface;
  IPCsObject = interface;
  IPCsObjectDisp = dispinterface;
  IPCsLayer = interface;
  IPCsLayerDisp = dispinterface;
  IPCsDrawing = interface;
  IPCsDrawingDisp = dispinterface;
  IPCsPages = interface;
  IPCsPagesDisp = dispinterface;
  IPCsPage = interface;
  IPCsPageDisp = dispinterface;
  IPCsSymbols = interface;
  IPCsSymbolsDisp = dispinterface;
  IPCsLibType = interface;
  IPCsLibTypeDisp = dispinterface;
  IPCsLibTypes = interface;
  IPCsLibTypesDisp = dispinterface;
  IPCsTexts = interface;
  IPCsTextsDisp = dispinterface;
  IPCsSymbolTexts = interface;
  IPCsSymbolTextsDisp = dispinterface;
  IPCsText = interface;
  IPCsTextDisp = dispinterface;
  IPCsTextFont = interface;
  IPCsTextFontDisp = dispinterface;
  IPCsNameText = interface;
  IPCsNameTextDisp = dispinterface;
  IPCsSymbol = interface;
  IPCsSymbolDisp = dispinterface;
  IPCsConnections = interface;
  IPCsConnectionsDisp = dispinterface;
  IPCsConnection = interface;
  IPCsConnectionDisp = dispinterface;
  IPCsConnectionTexts = interface;
  IPCsConnectionTextsDisp = dispinterface;
  IPCsDatafields = interface;
  IPCsDatafieldsDisp = dispinterface;
  IPCsDatafield = interface;
  IPCsDatafieldDisp = dispinterface;
  IPCsDataLink = interface;
  IPCsDataLinkDisp = dispinterface;
  IPCsPageSetup = interface;
  IPCsPageSetupDisp = dispinterface;
  IPCsPageReference = interface;
  IPCsPageReferenceDisp = dispinterface;
  IPCsBasePageRef = interface;
  IPCsBasePageRefDisp = dispinterface;
  IPCsExternalObjects = interface;
  IPCsExternalObjectsDisp = dispinterface;
  IPCsLayers = interface;
  IPCsLayersDisp = dispinterface;
  IPCsApplication = interface;
  IPCsApplicationDisp = dispinterface;
  IPCsDocuments = interface;
  IPCsDocumentsDisp = dispinterface;
  IPCsDocument = interface;
  IPCsDocumentDisp = dispinterface;
  IPCsApplicationPreferences = interface;
  IPCsApplicationPreferencesDisp = dispinterface;
  IPCsApplicationPreferences2 = interface;
  IPCsApplicationPreferences2Disp = dispinterface;
  IPCsSymbolMenu = interface;
  IPCsSymbolMenuDisp = dispinterface;
  IPCsComponentDatabase = interface;
  IPCsComponentDatabaseDisp = dispinterface;
  IPCsLibTypeConnections = interface;
  IPCsLibTypeConnectionsDisp = dispinterface;
  IPCsArcs = interface;
  IPCsArcsDisp = dispinterface;
  IPCsArc = interface;
  IPCsArcDisp = dispinterface;
  IApp = interface;
  IAppDisp = dispinterface;
  IPCsApplicationObsolete = interface;
  IPCsApplicationObsoleteDisp = dispinterface;
  IPCsApplicationEvents = dispinterface;
  IPCsApplication2 = interface;
  IPCsApplication2Disp = dispinterface;
  IPCsDocumentEvents = dispinterface;
  IPCsDocument2 = interface;
  IPCsDocument2Disp = dispinterface;
  IPCsApplicationPreferencesEvents = dispinterface;
  IPCsStatusPointEvents = dispinterface;
  IPCsLibType2 = interface;
  IPCsLibType2Disp = dispinterface;
  IPCsDrawingEvents = dispinterface;
  IPCsDrawing2 = interface;
  IPCsDrawing2Disp = dispinterface;
  IPCsLibTypesEvents = dispinterface;
  IPCsPagesEvents = dispinterface;
  IPCsPages2 = interface;
  IPCsPages2Disp = dispinterface;
  IPCsSymbol2 = interface;
  IPCsSymbol2Disp = dispinterface;
  IPCsSymbolsEnumerable = interface;
  IPCsSymbolsEnumerableDisp = dispinterface;
  IPCsLine2 = interface;
  IPCsLine2Disp = dispinterface;
  IPCsPageEvents = dispinterface;
  IPCsPage2 = interface;
  IPCsPage2Disp = dispinterface;
  IPCsArc2 = interface;
  IPCsArc2Disp = dispinterface;
  IPCsStatusObject = interface;
  IPCsStatusObjectDisp = dispinterface;
  IPCsConnection2 = interface;
  IPCsConnection2Disp = dispinterface;
  IPCsComponentDatabase2 = interface;
  IPCsComponentDatabase2Disp = dispinterface;
  IPCsGlobal = interface;
  IPCsGlobalDisp = dispinterface;
  IPCsGlobal2 = interface;
  IPCsGlobal2Disp = dispinterface;
  IPCsExternalObjects2 = interface;
  IPCsExternalObjects2Disp = dispinterface;
  IPCsDatafields2 = interface;
  IPCsDatafields2Disp = dispinterface;
  IPCsNameText2 = interface;
  IPCsNameText2Disp = dispinterface;
  IPCsReferenceDesignations = interface;
  IPCsReferenceDesignationsDisp = dispinterface;
  IPCsReferenceDesignations2 = interface;
  IPCsReferenceDesignations2Disp = dispinterface;
  IPCsReferenceDesignation = interface;
  IPCsReferenceDesignationDisp = dispinterface;
  IPCsSymbolFolderAliases = interface;
  IPCsSymbolFolderAliasesDisp = dispinterface;
  IPCsApplication3 = interface;
  IPCsApplication3Disp = dispinterface;
  IPCsDialog = interface;
  IPCsDialogDisp = dispinterface;
  IPCsDialogEvents = dispinterface;
  IPCsFileDialog = interface;
  IPCsFileDialogDisp = dispinterface;
  IPCsFileDialogEvents = dispinterface;
  IPCsLineTexts = interface;
  IPCsLineTextsDisp = dispinterface;
  IPCsLists = interface;
  IPCsListsDisp = dispinterface;
  IPCsLists2 = interface;
  IPCsLists2Disp = dispinterface;
  IPCsList = interface;
  IPCsListDisp = dispinterface;
  IPCsList2 = interface;
  IPCsList2Disp = dispinterface;
  IPCsExplorerWindow = interface;
  IPCsExplorerWindowDisp = dispinterface;
  IPCsExplorerWindowEvents = dispinterface;
  IPCsSubDrawing = interface;
  IPCsSubDrawingDisp = dispinterface;
  IPCsListConnData = interface;
  IPCsListConnDataDisp = dispinterface;
  IPCsSelectStatus = interface;
  IPCsSelectStatusDisp = dispinterface;
  IPCsSelectStatusEvents = dispinterface;
  IPCsSelectionArea = interface;
  IPCsSelectionAreaDisp = dispinterface;
  IPCsJoins = interface;
  IPCsJoinsDisp = dispinterface;
  IPCsJoin = interface;
  IPCsJoinDisp = dispinterface;
  IPCsDrawingPreferences = interface;
  IPCsDrawingPreferencesDisp = dispinterface;
  IPCsReference = interface;
  IPCsReferenceDisp = dispinterface;
  IPCsReferenceFrame = interface;
  IPCsReferenceFrameDisp = dispinterface;
  IPCsSymbol3 = interface;
  IPCsSymbol3Disp = dispinterface;
  IPCsLine3 = interface;
  IPCsLine3Disp = dispinterface;
  IPCsPage3 = interface;
  IPCsPage3Disp = dispinterface;
  IPCsDrawing3 = interface;
  IPCsDrawing3Disp = dispinterface;
  IPCsGlobal3 = interface;
  IPCsGlobal3Disp = dispinterface;
  IPCsDrawingPreferences2 = interface;
  IPCsDrawingPreferences2Disp = dispinterface;
  IPCsNameText3 = interface;
  IPCsNameText3Disp = dispinterface;
  IPCsSubDrawing2 = interface;
  IPCsSubDrawing2Disp = dispinterface;
  IPCsTemplateData = interface;
  IPCsTemplateDataDisp = dispinterface;
  IPCsExternalObject = interface;
  IPCsExternalObjectDisp = dispinterface;
  IPCsExternalObjectEvents = dispinterface;
  IPCsExternalObject2 = interface;
  IPCsExternalObject2Disp = dispinterface;
  IPCsPicture = interface;
  IPCsPictureDisp = dispinterface;
  IPCsPictureEvents = dispinterface;
  IPCsDatabaseSetup = interface;
  IPCsDatabaseSetupDisp = dispinterface;
  IPCsLine4 = interface;
  IPCsLine4Disp = dispinterface;
  IPCSLine5 = interface;
  IPCSLine5Disp = dispinterface;
  IPCsSymbol4 = interface;
  IPCsSymbol4Disp = dispinterface;
  IPCsPage4 = interface;
  IPCsPage4Disp = dispinterface;
  IPCsText2 = interface;
  IPCsText2Disp = dispinterface;
  IPCsNameText4 = interface;
  IPCsNameText4Disp = dispinterface;
  IPCsComponent = interface;
  IPCsComponentDisp = dispinterface;
  IPCsComponentEvents = dispinterface;
  IPCsComponent2 = interface;
  IPCsComponent2Disp = dispinterface;
  IPCsComponentName = interface;
  IPCsComponentNameDisp = dispinterface;
  IPCsObjectRefID = interface;
  IPCsObjectRefIDDisp = dispinterface;
  IPCsReferenceDesignations3 = interface;
  IPCsReferenceDesignations3Disp = dispinterface;
  IPCsReferenceDesignationList = interface;
  IPCsReferenceDesignationListDisp = dispinterface;
  IPCsLine6 = interface;
  IPCsLine6Disp = dispinterface;
  IPCsSymbolsEnumerator = interface;
  IPCsSymbolsEnumeratorDisp = dispinterface;
  IPCsComponentSymbols = interface;
  IPCsComponentSymbolsDisp = dispinterface;
  IPCsPage5 = interface;
  IPCsPage5Disp = dispinterface;
  IPCsLineDatafields = interface;
  IPCsLineDatafieldsDisp = dispinterface;
  IPCsRefIDFormat = interface;
  IPCsRefIDFormatDisp = dispinterface;
  IPCsCompName = interface;
  IPCsCompNameDisp = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library
// (NOTE: Here we map each CoClass to its Default Interface)
// *********************************************************************//
  PCsPoints = IPCsPoints;
  PCsDocuments = IPCsDocuments;
  PCsApplicationPreferences = IPCsApplicationPreferences;
  PCsStatusPoint = IPCsStatusPoint;
  PCsLibTypes = IPCsLibTypes;
  PCsLines = IPCsLines;
  PCsArcs = IPCsArcs;
  PCsPageSetup = IPCsPageSetup;
  PCsPageReference = IPCsPageReference;
  PCsBasePageRef = IPCsBasePageRef;
  PCsTexts = IPCsTexts;
  PCsSymbolTexts = IPCsSymbolTexts;
  PCsSymbolMenu = IPCsSymbolMenu;
  PCsLibTypeConnections = IPCsLibTypeConnections;
  PCsConnections = IPCsConnections;
  PCsConnectionTexts = IPCsConnectionTexts;
  PCsTextFont = IPCsTextFont;
  PCsJoinPoint = IPCsJoinPoint;
  PCsLayers = IPCsLayers;
  PCsLayer = IPCsLayer;
  PCsObject = IPCsObject;
  App = IApp;
  PCsDatafield = IPCsDatafield;
  PCsDataLink = IPCsDataLink;
  PCsReferenceDesignation = IPCsReferenceDesignation;
  PCsSymbolFolderAliases = IPCsSymbolFolderAliases;
  PCsDialog = IPCsDialog;
  PCsFileDialog = IPCsFileDialog;
  PCsLineTexts = IPCsLineTexts;
  PCsExplorerWindow = IPCsExplorerWindow;
  PCsSubDrawing = IPCsSubDrawing;
  PCsListConnData = IPCsListConnData;
  PCsSelectStatus = IPCsSelectStatus;
  PCsSelectionArea = IPCsSelectionArea;
  PCsJoins = IPCsJoins;
  PCsJoin = IPCsJoin;
  PCsDrawingPreferences = IPCsDrawingPreferences;
  PCsReference = IPCsReference;
  PCsReferenceFrame = IPCsReferenceFrame;
  PCsDocument = IPCsDocument;
  PCsFigure = IPCsFigure;
  PCsList = IPCsList;
  PCsLibType = IPCsLibType;
  PCsTemplateData = IPCsTemplateData;
  PCsPages = IPCsPages;
  PCsPicture = IPCsPicture;
  PCsDatabaseSetup = IPCsDatabaseSetup;
  Application = IPCsApplication;
  PCsDrawing = IPCsDrawing;
  PCsComponentDatabase = IPCsComponentDatabase;
  PCsLists = IPCsLists;
  PCsStatusObject = IPCsStatusObject;
  PCsConnection = IPCsConnection;
  PCsExternalObjects = IPCsExternalObjects;
  PCsExternalObject = IPCsExternalObject;
  PCsComponentName = IPCsComponentName;
  PCsSymbol = IPCsSymbol;
  PCsObjectRefID = IPCsObjectRefID;
  PCsReferenceDesignationList = IPCsReferenceDesignationList;
  PCsText = IPCsText;
  PCsArc = IPCsArc;
  PCsSymbols = IPCsSymbols;
  PCsComponentSymbols = IPCsComponentSymbols;
  PCsSymbolsEnumerator = IPCsSymbolsEnumerator;
  PCsLine = IPCsLine;
  PCsDatafields = IPCsDatafields;
  PCsLineDatafields = IPCsLineDatafields;
  PCsNameText = IPCsNameText;
  PCsRefIDFormat = IPCsRefIDFormat;
  PCsGlobal = IPCsGlobal;
  PCsReferenceDesignations = IPCsReferenceDesignations;
  PCsPage = IPCsPage;
  PCsComponent = IPCsComponent;
  PCsCompName = IPCsCompName;


// *********************************************************************//
// Declaration of structures, unions and aliases.
// *********************************************************************//
  POleVariant1 = ^OleVariant; {*}

  TWindowHandle = LongWord;
  PCsApplication = Application;
  TRGBColor = SYSINT;

  TPCsIOStatus = record
    MainType: psIOStatusMainType;
    Extension: psIOStatusExtension;
  end;

  TPCsPoint3D = record
    X: Integer;
    Y: Integer;
    Z: Integer;
  end;

  TPCsPoint2D = record
    X: Integer;
    Y: Integer;
  end;

  TPCsRect = record
    Left: Integer;
    Top: Integer;
    Right: Integer;
    Bottom: Integer;
  end;


// *********************************************************************//
// DispIntf:  IPCsFigureEvents
// Flags:     (0)
// GUID:      {281F4EB6-5696-4B49-ACEE-F3C0F1C4F6E3}
// *********************************************************************//
  IPCsFigureEvents = dispinterface
    ['{281F4EB6-5696-4B49-ACEE-F3C0F1C4F6E3}']
  end;

// *********************************************************************//
// Interface: IPCsFigure
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {381FE3A1-B600-41D8-99BB-611897416A12}
// *********************************************************************//
  IPCsFigure = interface(IDispatch)
    ['{381FE3A1-B600-41D8-99BB-611897416A12}']
    function Get_Lines: IPCsLines; safecall;
    function Get_Connections: IPCsLibTypeConnections; safecall;
    function Get_Arcs: IPCsArcs; safecall;
    function Get_Texts: IPCsTexts; safecall;
    function Get_Datafields: IPCsDatafields; safecall;
    function Get_Title: WideString; safecall;
    procedure Set_Title(const Value: WideString); safecall;
    function Get_TitleLC(LCID: Integer): WideString; safecall;
    procedure Set_TitleLC(LCID: Integer; const Value: WideString); safecall;
    function Get_States: OleVariant; safecall;
    function Get_SymbolTexts: IPCsSymbolTexts; safecall;
    property Lines: IPCsLines read Get_Lines;
    property Connections: IPCsLibTypeConnections read Get_Connections;
    property Arcs: IPCsArcs read Get_Arcs;
    property Texts: IPCsTexts read Get_Texts;
    property Datafields: IPCsDatafields read Get_Datafields;
    property Title: WideString read Get_Title write Set_Title;
    property TitleLC[LCID: Integer]: WideString read Get_TitleLC write Set_TitleLC;
    property States: OleVariant read Get_States;
    property SymbolTexts: IPCsSymbolTexts read Get_SymbolTexts;
  end;

// *********************************************************************//
// DispIntf:  IPCsFigureDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {381FE3A1-B600-41D8-99BB-611897416A12}
// *********************************************************************//
  IPCsFigureDisp = dispinterface
    ['{381FE3A1-B600-41D8-99BB-611897416A12}']
    property Lines: IPCsLines readonly dispid 201;
    property Connections: IPCsLibTypeConnections readonly dispid 202;
    property Arcs: IPCsArcs readonly dispid 203;
    property Texts: IPCsTexts readonly dispid 204;
    property Datafields: IPCsDatafields readonly dispid 205;
    property Title: WideString dispid 303;
    property TitleLC[LCID: Integer]: WideString dispid 206;
    property States: OleVariant readonly dispid 207;
    property SymbolTexts: IPCsSymbolTexts readonly dispid 208;
  end;

// *********************************************************************//
// Interface: IPCsBase
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {970CEB0E-ACAA-4E3F-B5B1-12833D73FF3B}
// *********************************************************************//
  IPCsBase = interface(IDispatch)
    ['{970CEB0E-ACAA-4E3F-B5B1-12833D73FF3B}']
    function Get_Handle: LongWord; safecall;
    function CustomCommand(const Command: WideString; Params: OleVariant; out Results: OleVariant): WordBool; safecall;
    function Equals(const Obj: IPCsBase): WordBool; safecall;
    procedure AssignTo(const Obj: IPCsBase); safecall;
    function Get_IsErased: WordBool; safecall;
    function Get_Application: IPCsApplication2; safecall;
    function Get_LStatus: Integer; safecall;
    property Handle: LongWord read Get_Handle;
    property IsErased: WordBool read Get_IsErased;
    property Application: IPCsApplication2 read Get_Application;
    property LStatus: Integer read Get_LStatus;
  end;

// *********************************************************************//
// DispIntf:  IPCsBaseDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {970CEB0E-ACAA-4E3F-B5B1-12833D73FF3B}
// *********************************************************************//
  IPCsBaseDisp = dispinterface
    ['{970CEB0E-ACAA-4E3F-B5B1-12833D73FF3B}']
    property Handle: LongWord readonly dispid 201;
    function CustomCommand(const Command: WideString; Params: OleVariant; out Results: OleVariant): WordBool; dispid 202;
    function Equals(const Obj: IPCsBase): WordBool; dispid 203;
    procedure AssignTo(const Obj: IPCsBase); dispid 204;
    property IsErased: WordBool readonly dispid 205;
    property Application: IPCsApplication2 readonly dispid 206;
    property LStatus: Integer readonly dispid 207;
  end;

// *********************************************************************//
// Interface: IPCsLines
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {DDE0BE45-1FBB-4729-AFB0-314BC3CEA689}
// *********************************************************************//
  IPCsLines = interface(IDispatch)
    ['{DDE0BE45-1FBB-4729-AFB0-314BC3CEA689}']
    function Add: IPCsLine; safecall;
    function Get__NewEnum: IEnumVARIANT; safecall;
    function Get_Count: Integer; safecall;
    function Get_Item(Index: Integer): IPCsLine; safecall;
    property _NewEnum: IEnumVARIANT read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Item[Index: Integer]: IPCsLine read Get_Item; default;
  end;

// *********************************************************************//
// DispIntf:  IPCsLinesDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {DDE0BE45-1FBB-4729-AFB0-314BC3CEA689}
// *********************************************************************//
  IPCsLinesDisp = dispinterface
    ['{DDE0BE45-1FBB-4729-AFB0-314BC3CEA689}']
    function Add: IPCsLine; dispid 201;
    property _NewEnum: IEnumVARIANT readonly dispid -4;
    property Count: Integer readonly dispid 203;
    property Item[Index: Integer]: IPCsLine readonly dispid 0; default;
  end;

// *********************************************************************//
// Interface: IPCsLine
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {13AECE11-B1EA-459C-BE19-7F484E12D3EE}
// *********************************************************************//
  IPCsLine = interface(IDispatch)
    ['{13AECE11-B1EA-459C-BE19-7F484E12D3EE}']
    function Get_Points: PCsPoints; safecall;
    procedure AddPoint(Pt: TPCsPoint3D); safecall;
    function Get_IsElectrical: WordBool; safecall;
    procedure Set_IsElectrical(Value: WordBool); safecall;
    function Get_LineStyle: PsLineStyle; safecall;
    procedure Set_LineStyle(Value: PsLineStyle); safecall;
    function Get_Color: PsPenColor; safecall;
    procedure Set_Color(Value: PsPenColor); safecall;
    function Get_PenWidth: Integer; safecall;
    procedure Set_PenWidth(Value: Integer); safecall;
    function Get_Quantity: Double; safecall;
    function Get_IsFilled: WordBool; safecall;
    procedure Set_IsFilled(Value: WordBool); safecall;
    function Get_IsJumperingLink: WordBool; safecall;
    procedure Set_IsJumperingLink(Value: WordBool); safecall;
    function Get_IsCableWire: WordBool; safecall;
    procedure Set_IsCableWire(Value: WordBool); safecall;
    function Get_IsPartLine: WordBool; safecall;
    function Get_Extended: WordBool; safecall;
    procedure Set_Extended(Value: WordBool); safecall;
    function Get_IsOpen: WordBool; safecall;
    procedure Set_IsOpen(Value: WordBool); safecall;
    function Get_IsSignalsJoined: WordBool; safecall;
    function AddBendingPoint(Pt: TPCsPoint3D; BendFactor: Double): WordBool; safecall;
    procedure ConnectTo(const JoinPoint: IPCsJoinPoint); safecall;
    function Get_Pattern: Integer; safecall;
    procedure Set_Pattern(Value: Integer); safecall;
    function Get_MultilineDistance: Integer; safecall;
    procedure Set_MultilineDistance(Value: Integer); safecall;
    function Get_Layer: IPCsLayer; safecall;
    procedure Set_Layer(const Value: IPCsLayer); safecall;
    function Get_IsCurvedLine: WordBool; safecall;
    procedure Set_IsCurvedLine(Value: WordBool); safecall;
    function Get_Visible: WordBool; safecall;
    procedure Set_Visible(Value: WordBool); safecall;
    procedure AddPointXYZ(X: Integer; Y: Integer; Z: Integer); safecall;
    procedure Delete; safecall;
    function Get_LineTexts: IPCsLineTexts; safecall;
    function Get_Handle: Integer; safecall;
    function Get_State: Integer; safecall;
    procedure Set_State(Value: Integer); safecall;
    function Get_ObjectType: PsPCsObjectType; safecall;
    procedure InitLineTexts; safecall;
    function MoveRelative(AOffset: TPCsPoint3D; const Figure: IPCsFigure): WordBool; safecall;
    function MoveRelativeXYZ(OffsetX: Integer; OffsetY: Integer; OffsetZ: Integer;
                             const Figure: IPCsFigure): WordBool; safecall;
    function ChangeDrawingOrder(AAction: psChangeDrawingOrder): WordBool; safecall;
    property Points: PCsPoints read Get_Points;
    property IsElectrical: WordBool read Get_IsElectrical write Set_IsElectrical;
    property LineStyle: PsLineStyle read Get_LineStyle write Set_LineStyle;
    property Color: PsPenColor read Get_Color write Set_Color;
    property PenWidth: Integer read Get_PenWidth write Set_PenWidth;
    property Quantity: Double read Get_Quantity;
    property IsFilled: WordBool read Get_IsFilled write Set_IsFilled;
    property IsJumperingLink: WordBool read Get_IsJumperingLink write Set_IsJumperingLink;
    property IsCableWire: WordBool read Get_IsCableWire write Set_IsCableWire;
    property IsPartLine: WordBool read Get_IsPartLine;
    property Extended: WordBool read Get_Extended write Set_Extended;
    property IsOpen: WordBool read Get_IsOpen write Set_IsOpen;
    property IsSignalsJoined: WordBool read Get_IsSignalsJoined;
    property Pattern: Integer read Get_Pattern write Set_Pattern;
    property MultilineDistance: Integer read Get_MultilineDistance write Set_MultilineDistance;
    property Layer: IPCsLayer read Get_Layer write Set_Layer;
    property IsCurvedLine: WordBool read Get_IsCurvedLine write Set_IsCurvedLine;
    property Visible: WordBool read Get_Visible write Set_Visible;
    property LineTexts: IPCsLineTexts read Get_LineTexts;
    property Handle: Integer read Get_Handle;
    property State: Integer read Get_State write Set_State;
    property ObjectType: PsPCsObjectType read Get_ObjectType;
  end;

// *********************************************************************//
// DispIntf:  IPCsLineDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {13AECE11-B1EA-459C-BE19-7F484E12D3EE}
// *********************************************************************//
  IPCsLineDisp = dispinterface
    ['{13AECE11-B1EA-459C-BE19-7F484E12D3EE}']
    property Points: PCsPoints readonly dispid 401;
    procedure AddPoint(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 402;
    property IsElectrical: WordBool dispid 301;
    property LineStyle: PsLineStyle dispid 302;
    property Color: PsPenColor dispid 303;
    property PenWidth: Integer dispid 304;
    property Quantity: Double readonly dispid 305;
    property IsFilled: WordBool dispid 306;
    property IsJumperingLink: WordBool dispid 308;
    property IsCableWire: WordBool dispid 307;
    property IsPartLine: WordBool readonly dispid 310;
    property Extended: WordBool dispid 311;
    property IsOpen: WordBool dispid 312;
    property IsSignalsJoined: WordBool readonly dispid 313;
    function AddBendingPoint(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; BendFactor: Double): WordBool; dispid 315;
    procedure ConnectTo(const JoinPoint: IPCsJoinPoint); dispid 316;
    property Pattern: Integer dispid 317;
    property MultilineDistance: Integer dispid 318;
    property Layer: IPCsLayer dispid 319;
    property IsCurvedLine: WordBool dispid 309;
    property Visible: WordBool dispid 201;
    procedure AddPointXYZ(X: Integer; Y: Integer; Z: Integer); dispid 202;
    procedure Delete; dispid 203;
    property LineTexts: IPCsLineTexts readonly dispid 204;
    property Handle: Integer readonly dispid 205;
    property State: Integer dispid 250;
    property ObjectType: PsPCsObjectType readonly dispid 450;
    procedure InitLineTexts; dispid 206;
    function MoveRelative(AOffset: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Figure: IPCsFigure): WordBool; dispid 207;
    function MoveRelativeXYZ(OffsetX: Integer; OffsetY: Integer; OffsetZ: Integer;
                             const Figure: IPCsFigure): WordBool; dispid 208;
    function ChangeDrawingOrder(AAction: psChangeDrawingOrder): WordBool; dispid 209;
  end;

// *********************************************************************//
// Interface: IPCsPoints
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {647664A1-8E29-4469-ACE8-26608477DF8C}
// *********************************************************************//
  IPCsPoints = interface(IDispatch)
    ['{647664A1-8E29-4469-ACE8-26608477DF8C}']
    function Get_Item(Index: Integer): IPCsStatusPoint; safecall;
    function Get__NewEnum: IEnumVARIANT; safecall;
    function Get_Count: Integer; safecall;
    function Add: IPCsStatusPoint; safecall;
    property Item[Index: Integer]: IPCsStatusPoint read Get_Item; default;
    property _NewEnum: IEnumVARIANT read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  IPCsPointsDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {647664A1-8E29-4469-ACE8-26608477DF8C}
// *********************************************************************//
  IPCsPointsDisp = dispinterface
    ['{647664A1-8E29-4469-ACE8-26608477DF8C}']
    property Item[Index: Integer]: IPCsStatusPoint readonly dispid 0; default;
    property _NewEnum: IEnumVARIANT readonly dispid -4;
    property Count: Integer readonly dispid 201;
    function Add: IPCsStatusPoint; dispid 202;
  end;

// *********************************************************************//
// Interface: IPCsStatusPoint
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {28DF9ED4-686A-4317-8843-4EFAE1CD3C5F}
// *********************************************************************//
  IPCsStatusPoint = interface(IDispatch)
    ['{28DF9ED4-686A-4317-8843-4EFAE1CD3C5F}']
    function Get_Visible: WordBool; safecall;
    procedure Set_Visible(Value: WordBool); safecall;
    function Get_Position: TPCsPoint3D; safecall;
    procedure Set_Position(Value: TPCsPoint3D); safecall;
    function Get_LStatus: Integer; safecall;
    procedure Set_LStatus(Value: Integer); safecall;
    function Move(Pt: TPCsPoint3D; const Figure: IPCsFigure): WordBool; safecall;
    procedure Rotate(DeltaAngle: Integer; Pt: TPCsPoint3D); safecall;
    procedure Scale(Factor: Single; Pt: TPCsPoint3D); safecall;
    function Get_Handle: LongWord; safecall;
    procedure GetPosition(out X: Integer; out Y: Integer; out Z: Integer); safecall;
    procedure SetPosition(X: Integer; Y: Integer; Z: Integer); safecall;
    function Get_State: Integer; safecall;
    procedure Set_State(Value: Integer); safecall;
    function ChangeDrawingOrder(AAction: psChangeDrawingOrder): WordBool; safecall;
    property Visible: WordBool read Get_Visible write Set_Visible;
    property Position: TPCsPoint3D read Get_Position write Set_Position;
    property LStatus: Integer read Get_LStatus write Set_LStatus;
    property Handle: LongWord read Get_Handle;
    property State: Integer read Get_State write Set_State;
  end;

// *********************************************************************//
// DispIntf:  IPCsStatusPointDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {28DF9ED4-686A-4317-8843-4EFAE1CD3C5F}
// *********************************************************************//
  IPCsStatusPointDisp = dispinterface
    ['{28DF9ED4-686A-4317-8843-4EFAE1CD3C5F}']
    property Visible: WordBool dispid 201;
    property Position: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant dispid 202;
    property LStatus: Integer dispid 203;
    function Move(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Figure: IPCsFigure): WordBool; dispid 204;
    procedure Rotate(DeltaAngle: Integer; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 205;
    procedure Scale(Factor: Single; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 208;
    property Handle: LongWord readonly dispid 206;
    procedure GetPosition(out X: Integer; out Y: Integer; out Z: Integer); dispid 207;
    procedure SetPosition(X: Integer; Y: Integer; Z: Integer); dispid 209;
    property State: Integer dispid 250;
    function ChangeDrawingOrder(AAction: psChangeDrawingOrder): WordBool; dispid 210;
  end;

// *********************************************************************//
// Interface: IPCsJoinPoint
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {3FB4759B-B7DE-4801-9375-6872D0B4A4A8}
// *********************************************************************//
  IPCsJoinPoint = interface(IPCsStatusPoint)
    ['{3FB4759B-B7DE-4801-9375-6872D0B4A4A8}']
  end;

// *********************************************************************//
// DispIntf:  IPCsJoinPointDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {3FB4759B-B7DE-4801-9375-6872D0B4A4A8}
// *********************************************************************//
  IPCsJoinPointDisp = dispinterface
    ['{3FB4759B-B7DE-4801-9375-6872D0B4A4A8}']
    property Visible: WordBool dispid 201;
    property Position: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant dispid 202;
    property LStatus: Integer dispid 203;
    function Move(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Figure: IPCsFigure): WordBool; dispid 204;
    procedure Rotate(DeltaAngle: Integer; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 205;
    procedure Scale(Factor: Single; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 208;
    property Handle: LongWord readonly dispid 206;
    procedure GetPosition(out X: Integer; out Y: Integer; out Z: Integer); dispid 207;
    procedure SetPosition(X: Integer; Y: Integer; Z: Integer); dispid 209;
    property State: Integer dispid 250;
    function ChangeDrawingOrder(AAction: psChangeDrawingOrder): WordBool; dispid 210;
  end;

// *********************************************************************//
// Interface: IPCsObject
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0896A10E-AF2C-450B-ACC0-32B4D0FD52C6}
// *********************************************************************//
  IPCsObject = interface(IDispatch)
    ['{0896A10E-AF2C-450B-ACC0-32B4D0FD52C6}']
    function HasSameOwnerDrawing(const AObject: IPCsObject): WordBool; safecall;
    function Get_OwnerDrawing: IPCsDrawing; safecall;
    property OwnerDrawing: IPCsDrawing read Get_OwnerDrawing;
  end;

// *********************************************************************//
// DispIntf:  IPCsObjectDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0896A10E-AF2C-450B-ACC0-32B4D0FD52C6}
// *********************************************************************//
  IPCsObjectDisp = dispinterface
    ['{0896A10E-AF2C-450B-ACC0-32B4D0FD52C6}']
    function HasSameOwnerDrawing(const AObject: IPCsObject): WordBool; dispid 190;
    property OwnerDrawing: IPCsDrawing readonly dispid 191;
  end;

// *********************************************************************//
// Interface: IPCsLayer
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {60A0B67B-DCBD-4ADE-8E39-12761C092AAC}
// *********************************************************************//
  IPCsLayer = interface(IPCsObject)
    ['{60A0B67B-DCBD-4ADE-8E39-12761C092AAC}']
    function Get_Index: Integer; safecall;
    function Get_Name: WideString; safecall;
    procedure Set_Name(const Value: WideString); safecall;
    function Get_Visible: WordBool; safecall;
    procedure Set_Visible(Value: WordBool); safecall;
    function Get_IsUsed: WordBool; safecall;
    property Index: Integer read Get_Index;
    property Name: WideString read Get_Name write Set_Name;
    property Visible: WordBool read Get_Visible write Set_Visible;
    property IsUsed: WordBool read Get_IsUsed;
  end;

// *********************************************************************//
// DispIntf:  IPCsLayerDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {60A0B67B-DCBD-4ADE-8E39-12761C092AAC}
// *********************************************************************//
  IPCsLayerDisp = dispinterface
    ['{60A0B67B-DCBD-4ADE-8E39-12761C092AAC}']
    property Index: Integer readonly dispid 201;
    property Name: WideString dispid 202;
    property Visible: WordBool dispid 203;
    property IsUsed: WordBool readonly dispid 301;
    function HasSameOwnerDrawing(const AObject: IPCsObject): WordBool; dispid 190;
    property OwnerDrawing: IPCsDrawing readonly dispid 191;
  end;

// *********************************************************************//
// Interface: IPCsDrawing
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {61BFFC36-9B1E-4BFC-92ED-EB532CD10268}
// *********************************************************************//
  IPCsDrawing = interface(IDispatch)
    ['{61BFFC36-9B1E-4BFC-92ED-EB532CD10268}']
    function Get_Pages: IPCsPages; safecall;
    function Get_LibTypes: IPCsLibTypes; safecall;
    function Get_Layers: IPCsLayers; safecall;
    function Get_Application: IPCsApplication; safecall;
    function Get_Parent: IPCsDocument; safecall;
    function Get_Title: WideString; safecall;
    procedure Set_Title(const Value: WideString); safecall;
    function GetSymbolFromHandle(PageNo: Integer; SymbolHandle: Integer): IPCsSymbol; safecall;
    function Get_ActiveLayer: IPCsLayer; safecall;
    procedure Set_ActiveLayer(const Value: IPCsLayer); safecall;
    function Get_ProjectData(Index: Integer): WideString; safecall;
    procedure Set_ProjectData(Index: Integer; const Value: WideString); safecall;
    function Get_ProjectDataDefs: OleVariant; safecall;
    procedure Set_ProjectDataDefs(Value: OleVariant); safecall;
    function Get_PageDataDefs: OleVariant; safecall;
    procedure Set_PageDataDefs(Value: OleVariant); safecall;
    function Get_SymbolDataDefs: OleVariant; safecall;
    procedure Set_SymbolDataDefs(Value: OleVariant); safecall;
    function Get_ReferenceDesignations: IPCsReferenceDesignations; safecall;
    function SetProjectData(const FieldName: WideString; const Value: WideString): Integer; safecall;
    function Get_Lists: IPCsLists; safecall;
    procedure OptimizeLibtypes; safecall;
    function LoadSubDrawing(const Filename: WideString): IPCsSubDrawing; safecall;
    function GetLineFromHandle(PageNo: Integer; LineHandle: Integer): IPCsLine; safecall;
    function GetTextFromHandle(PageNo: Integer; TextHandle: Integer): IPCsText; safecall;
    function Get_SystemVariable(const Variable: WideString): OleVariant; safecall;
    procedure Set_SystemVariable(const Variable: WideString; Value: OleVariant); safecall;
    function Get_Preferences: IPCsDrawingPreferences; safecall;
    procedure SynchronizePLCData(const ParamStr: WideString); safecall;
    procedure ApplyIntelligentInvisibility(const ParamStr: WideString); safecall;
    function UpdateFromDatabase(Reserved: Integer): WordBool; safecall;
    function Get_Document: IPCsDocument; safecall;
    property Pages: IPCsPages read Get_Pages;
    property LibTypes: IPCsLibTypes read Get_LibTypes;
    property Layers: IPCsLayers read Get_Layers;
    property Application: IPCsApplication read Get_Application;
    property Parent: IPCsDocument read Get_Parent;
    property Title: WideString read Get_Title write Set_Title;
    property ActiveLayer: IPCsLayer read Get_ActiveLayer write Set_ActiveLayer;
    property ProjectData[Index: Integer]: WideString read Get_ProjectData write Set_ProjectData;
    property ProjectDataDefs: OleVariant read Get_ProjectDataDefs write Set_ProjectDataDefs;
    property PageDataDefs: OleVariant read Get_PageDataDefs write Set_PageDataDefs;
    property SymbolDataDefs: OleVariant read Get_SymbolDataDefs write Set_SymbolDataDefs;
    property ReferenceDesignations: IPCsReferenceDesignations read Get_ReferenceDesignations;
    property Lists: IPCsLists read Get_Lists;
    property SystemVariable[const Variable: WideString]: OleVariant read Get_SystemVariable write Set_SystemVariable;
    property Preferences: IPCsDrawingPreferences read Get_Preferences;
    property Document: IPCsDocument read Get_Document;
  end;

// *********************************************************************//
// DispIntf:  IPCsDrawingDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {61BFFC36-9B1E-4BFC-92ED-EB532CD10268}
// *********************************************************************//
  IPCsDrawingDisp = dispinterface
    ['{61BFFC36-9B1E-4BFC-92ED-EB532CD10268}']
    property Pages: IPCsPages readonly dispid 201;
    property LibTypes: IPCsLibTypes readonly dispid 202;
    property Layers: IPCsLayers readonly dispid 205;
    property Application: IPCsApplication readonly dispid 206;
    property Parent: IPCsDocument readonly dispid 207;
    property Title: WideString dispid 203;
    function GetSymbolFromHandle(PageNo: Integer; SymbolHandle: Integer): IPCsSymbol; dispid 204;
    property ActiveLayer: IPCsLayer dispid 208;
    property ProjectData[Index: Integer]: WideString dispid 209;
    property ProjectDataDefs: OleVariant dispid 210;
    property PageDataDefs: OleVariant dispid 211;
    property SymbolDataDefs: OleVariant dispid 214;
    property ReferenceDesignations: IPCsReferenceDesignations readonly dispid 212;
    function SetProjectData(const FieldName: WideString; const Value: WideString): Integer; dispid 213;
    property Lists: IPCsLists readonly dispid 215;
    procedure OptimizeLibtypes; dispid 216;
    function LoadSubDrawing(const Filename: WideString): IPCsSubDrawing; dispid 217;
    function GetLineFromHandle(PageNo: Integer; LineHandle: Integer): IPCsLine; dispid 218;
    function GetTextFromHandle(PageNo: Integer; TextHandle: Integer): IPCsText; dispid 219;
    property SystemVariable[const Variable: WideString]: OleVariant dispid 220;
    property Preferences: IPCsDrawingPreferences readonly dispid 221;
    procedure SynchronizePLCData(const ParamStr: WideString); dispid 222;
    procedure ApplyIntelligentInvisibility(const ParamStr: WideString); dispid 223;
    function UpdateFromDatabase(Reserved: Integer): WordBool; dispid 224;
    property Document: IPCsDocument readonly dispid 225;
  end;

// *********************************************************************//
// Interface: IPCsPages
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {DAC4E978-E186-4E2F-925F-E796EC3442D6}
// *********************************************************************//
  IPCsPages = interface(IDispatch)
    ['{DAC4E978-E186-4E2F-925F-E796EC3442D6}']
    function Get_Item(Index: Integer): IPCsPage; safecall;
    function Get_Count: Integer; safecall;
    function Get__NewEnum: IEnumVARIANT; safecall;
    function Add(PageFunction: psPageFunction; const Template: WideString): IPCsPage; safecall;
    function Insert(Index: Integer; PageFunction: psPageFunction; const Template: WideString): IPCsPage; safecall;
    procedure RemovePages(IndexFrom: Integer; IndexTo: Integer); safecall;
    property Item[Index: Integer]: IPCsPage read Get_Item; default;
    property Count: Integer read Get_Count;
    property _NewEnum: IEnumVARIANT read Get__NewEnum;
  end;

// *********************************************************************//
// DispIntf:  IPCsPagesDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {DAC4E978-E186-4E2F-925F-E796EC3442D6}
// *********************************************************************//
  IPCsPagesDisp = dispinterface
    ['{DAC4E978-E186-4E2F-925F-E796EC3442D6}']
    property Item[Index: Integer]: IPCsPage readonly dispid 0; default;
    property Count: Integer readonly dispid 203;
    property _NewEnum: IEnumVARIANT readonly dispid -4;
    function Add(PageFunction: psPageFunction; const Template: WideString): IPCsPage; dispid 204;
    function Insert(Index: Integer; PageFunction: psPageFunction; const Template: WideString): IPCsPage; dispid 201;
    procedure RemovePages(IndexFrom: Integer; IndexTo: Integer); dispid 202;
  end;

// *********************************************************************//
// Interface: IPCsPage
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {5771AA42-30B5-496E-AB85-FB3CCE6F69B4}
// *********************************************************************//
  IPCsPage = interface(IPCsFigure)
    ['{5771AA42-30B5-496E-AB85-FB3CCE6F69B4}']
    function Get_Symbols: IPCsSymbols; safecall;
    function Get_PageNumber: WideString; safecall;
    procedure Set_PageNumber(const Value: WideString); safecall;
    function Get_PageSetup: IPCsPageSetup; safecall;
    function Get_PageHeader: IPCsSymbol; safecall;
    procedure Set_PageHeader(const Value: IPCsSymbol); safecall;
    function LoadPageHeader(const Filename: WideString): WordBool; safecall;
    function Get_IsEmpty: WordBool; safecall;
    function Get_PageType: psPageType; safecall;
    procedure Set_PageType(Value: psPageType); safecall;
    function Get_PageFunction: psPageFunction; safecall;
    procedure Set_PageFunction(Value: psPageFunction); safecall;
    function Get_Origo: TPCsPoint3D; safecall;
    procedure Set_Origo(Value: TPCsPoint3D); safecall;
    function Get_Index: Integer; safecall;
    procedure Set_Index(Value: Integer); safecall;
    procedure CreateJoins; safecall;
    procedure JoinParkedLines; safecall;
    procedure JoinConnectionsToCrossingLines; safecall;
    function PlaceSymbol(Pt: TPCsPoint3D; const Filename: WideString): IPCsSymbol; safecall;
    function Get_PageReference: IPCsPageReference; safecall;
    function Get_Scale: Single; safecall;
    procedure Set_Scale(Value: Single); safecall;
    function GetMasterPage: IPCsPage; safecall;
    function Get_ReadingDirection: Integer; safecall;
    procedure Set_ReadingDirection(Value: Integer); safecall;
    procedure Draw(DeviceContext: LongWord; Rect: TPCsRect); safecall;
    procedure DrawOntoWindow(Wnd: TWindowHandle; Rect: TPCsRect); safecall;
    function Get_ExternalObjects: IPCsExternalObjects; safecall;
    function GetRect(AFlags: psPageRectFlag): TPCsRect; safecall;
    function Get_PageData(Index: Integer): WideString; safecall;
    procedure Set_PageData(Index: Integer; const Value: WideString); safecall;
    function Get_PageDataCount: Integer; safecall;
    procedure Delete; safecall;
    procedure PrintOut; safecall;
    procedure Move(Index: Integer); safecall;
    function GetVirtualsAsSymbols: IPCsSymbols; safecall;
    function PlaceSubDrawing(const SubDrawing: IPCsSubDrawing; Group: Integer; X: Integer;
                             Y: Integer; Z: Integer; const Layer: IPCsLayer;
                             Options: psPlaceSubDrawingOption): WordBool; safecall;
    procedure SetPageHeaderLibtype(const LibType: IPCsLibType); safecall;
    function Get_Joins: IPCsJoins; safecall;
    procedure OffsetMoveObjects(AX: Integer; AY: Integer; AZ: Integer); safecall;
    function Get_RefDesLocation: WideString; safecall;
    procedure Set_RefDesLocation(const Value: WideString); safecall;
    function Get_RefDesFunction: WideString; safecall;
    procedure Set_RefDesFunction(const Value: WideString); safecall;
    function UpdateFromDatabase(Reserved: Integer): WordBool; safecall;
    property Symbols: IPCsSymbols read Get_Symbols;
    property PageNumber: WideString read Get_PageNumber write Set_PageNumber;
    property PageSetup: IPCsPageSetup read Get_PageSetup;
    property PageHeader: IPCsSymbol read Get_PageHeader write Set_PageHeader;
    property IsEmpty: WordBool read Get_IsEmpty;
    property PageType: psPageType read Get_PageType write Set_PageType;
    property PageFunction: psPageFunction read Get_PageFunction write Set_PageFunction;
    property Origo: TPCsPoint3D read Get_Origo write Set_Origo;
    property Index: Integer read Get_Index write Set_Index;
    property PageReference: IPCsPageReference read Get_PageReference;
    property Scale: Single read Get_Scale write Set_Scale;
    property ReadingDirection: Integer read Get_ReadingDirection write Set_ReadingDirection;
    property ExternalObjects: IPCsExternalObjects read Get_ExternalObjects;
    property PageData[Index: Integer]: WideString read Get_PageData write Set_PageData;
    property PageDataCount: Integer read Get_PageDataCount;
    property Joins: IPCsJoins read Get_Joins;
    property RefDesLocation: WideString read Get_RefDesLocation write Set_RefDesLocation;
    property RefDesFunction: WideString read Get_RefDesFunction write Set_RefDesFunction;
  end;

// *********************************************************************//
// DispIntf:  IPCsPageDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {5771AA42-30B5-496E-AB85-FB3CCE6F69B4}
// *********************************************************************//
  IPCsPageDisp = dispinterface
    ['{5771AA42-30B5-496E-AB85-FB3CCE6F69B4}']
    property Symbols: IPCsSymbols readonly dispid 401;
    property PageNumber: WideString dispid 402;
    property PageSetup: IPCsPageSetup readonly dispid 403;
    property PageHeader: IPCsSymbol dispid 404;
    function LoadPageHeader(const Filename: WideString): WordBool; dispid 405;
    property IsEmpty: WordBool readonly dispid 406;
    property PageType: psPageType dispid 407;
    property PageFunction: psPageFunction dispid 408;
    property Origo: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant dispid 409;
    property Index: Integer dispid 410;
    procedure CreateJoins; dispid 412;
    procedure JoinParkedLines; dispid 413;
    procedure JoinConnectionsToCrossingLines; dispid 414;
    function PlaceSymbol(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Filename: WideString): IPCsSymbol; dispid 415;
    property PageReference: IPCsPageReference readonly dispid 416;
    property Scale: Single dispid 301;
    function GetMasterPage: IPCsPage; dispid 311;
    property ReadingDirection: Integer dispid 304;
    procedure Draw(DeviceContext: LongWord; Rect: {NOT_OLEAUTO(TPCsRect)}OleVariant); dispid 305;
    procedure DrawOntoWindow(Wnd: TWindowHandle; Rect: {NOT_OLEAUTO(TPCsRect)}OleVariant); dispid 306;
    property ExternalObjects: IPCsExternalObjects readonly dispid 307;
    function GetRect(AFlags: psPageRectFlag): {NOT_OLEAUTO(TPCsRect)}OleVariant; dispid 302;
    property PageData[Index: Integer]: WideString dispid 308;
    property PageDataCount: Integer readonly dispid 309;
    procedure Delete; dispid 310;
    procedure PrintOut; dispid 312;
    procedure Move(Index: Integer); dispid 313;
    function GetVirtualsAsSymbols: IPCsSymbols; dispid 314;
    function PlaceSubDrawing(const SubDrawing: IPCsSubDrawing; Group: Integer; X: Integer;
                             Y: Integer; Z: Integer; const Layer: IPCsLayer;
                             Options: psPlaceSubDrawingOption): WordBool; dispid 315;
    procedure SetPageHeaderLibtype(const LibType: IPCsLibType); dispid 316;
    property Joins: IPCsJoins readonly dispid 500;
    procedure OffsetMoveObjects(AX: Integer; AY: Integer; AZ: Integer); dispid 317;
    property RefDesLocation: WideString dispid 318;
    property RefDesFunction: WideString dispid 319;
    function UpdateFromDatabase(Reserved: Integer): WordBool; dispid 320;
    property Lines: IPCsLines readonly dispid 201;
    property Connections: IPCsLibTypeConnections readonly dispid 202;
    property Arcs: IPCsArcs readonly dispid 203;
    property Texts: IPCsTexts readonly dispid 204;
    property Datafields: IPCsDatafields readonly dispid 205;
    property Title: WideString dispid 303;
    property TitleLC[LCID: Integer]: WideString dispid 206;
    property States: OleVariant readonly dispid 207;
    property SymbolTexts: IPCsSymbolTexts readonly dispid 208;
  end;

// *********************************************************************//
// Interface: IPCsSymbols
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {FF44C293-724F-4E91-AA14-17825A792DEC}
// *********************************************************************//
  IPCsSymbols = interface(IDispatch)
    ['{FF44C293-724F-4E91-AA14-17825A792DEC}']
    function Add(const LibType: IPCsLibType): IPCsSymbol; safecall;
    function Get_Item(Index: Integer): IPCsSymbol; safecall;
    function Get_Count: Integer; safecall;
    function Get__NewEnum: IEnumVARIANT; safecall;
    property Item[Index: Integer]: IPCsSymbol read Get_Item; default;
    property Count: Integer read Get_Count;
    property _NewEnum: IEnumVARIANT read Get__NewEnum;
  end;

// *********************************************************************//
// DispIntf:  IPCsSymbolsDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {FF44C293-724F-4E91-AA14-17825A792DEC}
// *********************************************************************//
  IPCsSymbolsDisp = dispinterface
    ['{FF44C293-724F-4E91-AA14-17825A792DEC}']
    function Add(const LibType: IPCsLibType): IPCsSymbol; dispid 400;
    property Item[Index: Integer]: IPCsSymbol readonly dispid 0; default;
    property Count: Integer readonly dispid 402;
    property _NewEnum: IEnumVARIANT readonly dispid -4;
  end;

// *********************************************************************//
// Interface: IPCsLibType
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {EF5253F8-0F9C-47A5-8BA7-2D369B12A62D}
// *********************************************************************//
  IPCsLibType = interface(IPCsFigure)
    ['{EF5253F8-0F9C-47A5-8BA7-2D369B12A62D}']
    function Get_SymbolType: psSymbolType; safecall;
    procedure Set_SymbolType(Value: psSymbolType); safecall;
    function Get_Name: WideString; safecall;
    procedure Set_Name(const Value: WideString); safecall;
    function Get_FileTime: Integer; safecall;
    procedure Set_FileTime(Value: Integer); safecall;
    function Get_IsReferenceOnly: WordBool; safecall;
    function Get_Handle: LongWord; safecall;
    function Get_Parent: IPCsLibTypes; safecall;
    function GetRect(AType: PsSymbolRectType): TPCsRect; safecall;
    function Get_IsElectrical: WordBool; safecall;
    procedure Set_IsElectrical(Value: WordBool); safecall;
    function Get_IsMechanical: WordBool; safecall;
    procedure Set_IsMechanical(Value: WordBool); safecall;
    procedure Delete; safecall;
    function Get_IsUsed: WordBool; safecall;
    function Get_ObjectType: PsPCsObjectType; safecall;
    property SymbolType: psSymbolType read Get_SymbolType write Set_SymbolType;
    property Name: WideString read Get_Name write Set_Name;
    property FileTime: Integer read Get_FileTime write Set_FileTime;
    property IsReferenceOnly: WordBool read Get_IsReferenceOnly;
    property Handle: LongWord read Get_Handle;
    property Parent: IPCsLibTypes read Get_Parent;
    property IsElectrical: WordBool read Get_IsElectrical write Set_IsElectrical;
    property IsMechanical: WordBool read Get_IsMechanical write Set_IsMechanical;
    property IsUsed: WordBool read Get_IsUsed;
    property ObjectType: PsPCsObjectType read Get_ObjectType;
  end;

// *********************************************************************//
// DispIntf:  IPCsLibTypeDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {EF5253F8-0F9C-47A5-8BA7-2D369B12A62D}
// *********************************************************************//
  IPCsLibTypeDisp = dispinterface
    ['{EF5253F8-0F9C-47A5-8BA7-2D369B12A62D}']
    property SymbolType: psSymbolType dispid 302;
    property Name: WideString dispid 304;
    property FileTime: Integer dispid 305;
    property IsReferenceOnly: WordBool readonly dispid 306;
    property Handle: LongWord readonly dispid 307;
    property Parent: IPCsLibTypes readonly dispid 301;
    function GetRect(AType: PsSymbolRectType): {NOT_OLEAUTO(TPCsRect)}OleVariant; dispid 309;
    property IsElectrical: WordBool dispid 310;
    property IsMechanical: WordBool dispid 311;
    procedure Delete; dispid 312;
    property IsUsed: WordBool readonly dispid 313;
    property ObjectType: PsPCsObjectType readonly dispid 450;
    property Lines: IPCsLines readonly dispid 201;
    property Connections: IPCsLibTypeConnections readonly dispid 202;
    property Arcs: IPCsArcs readonly dispid 203;
    property Texts: IPCsTexts readonly dispid 204;
    property Datafields: IPCsDatafields readonly dispid 205;
    property Title: WideString dispid 303;
    property TitleLC[LCID: Integer]: WideString dispid 206;
    property States: OleVariant readonly dispid 207;
    property SymbolTexts: IPCsSymbolTexts readonly dispid 208;
  end;

// *********************************************************************//
// Interface: IPCsLibTypes
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {4C394DA5-F3F7-49AE-BFA0-84899BBB6D25}
// *********************************************************************//
  IPCsLibTypes = interface(IDispatch)
    ['{4C394DA5-F3F7-49AE-BFA0-84899BBB6D25}']
    function Get_Item(Index: Integer): IPCsLibType; safecall;
    function Add(const Name: WideString): IPCsLibType; safecall;
    function Find(const Name: WideString; Time: Integer): IPCsLibType; safecall;
    function Load(const Filename: WideString): IPCsLibType; safecall;
    function Get_Count: Integer; safecall;
    function Get__NewEnum: IEnumVARIANT; safecall;
    function GetHandle: Integer; safecall;
    property Item[Index: Integer]: IPCsLibType read Get_Item; default;
    property Count: Integer read Get_Count;
    property _NewEnum: IEnumVARIANT read Get__NewEnum;
  end;

// *********************************************************************//
// DispIntf:  IPCsLibTypesDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {4C394DA5-F3F7-49AE-BFA0-84899BBB6D25}
// *********************************************************************//
  IPCsLibTypesDisp = dispinterface
    ['{4C394DA5-F3F7-49AE-BFA0-84899BBB6D25}']
    property Item[Index: Integer]: IPCsLibType readonly dispid 0; default;
    function Add(const Name: WideString): IPCsLibType; dispid 202;
    function Find(const Name: WideString; Time: Integer): IPCsLibType; dispid 203;
    function Load(const Filename: WideString): IPCsLibType; dispid 204;
    property Count: Integer readonly dispid 205;
    property _NewEnum: IEnumVARIANT readonly dispid -4;
    function GetHandle: Integer; dispid 201;
  end;

// *********************************************************************//
// Interface: IPCsTexts
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {81F82835-4133-46BA-8547-00BDB5D73C8F}
// *********************************************************************//
  IPCsTexts = interface(IDispatch)
    ['{81F82835-4133-46BA-8547-00BDB5D73C8F}']
    function Get__NewEnum: IEnumVARIANT; safecall;
    function Get_Item(Index: Integer): IPCsText; safecall;
    function Get_Count: Integer; safecall;
    function Add(Pt: TPCsPoint3D; const Value: WideString): IPCsText; safecall;
    function FindByValue(const Value: WideString): IPCsText; safecall;
    procedure Remove(Index: Integer); safecall;
    function AddXYZ(X: Integer; Y: Integer; Z: Integer; const Value: WideString): IPCsText; safecall;
    property _NewEnum: IEnumVARIANT read Get__NewEnum;
    property Item[Index: Integer]: IPCsText read Get_Item; default;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  IPCsTextsDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {81F82835-4133-46BA-8547-00BDB5D73C8F}
// *********************************************************************//
  IPCsTextsDisp = dispinterface
    ['{81F82835-4133-46BA-8547-00BDB5D73C8F}']
    property _NewEnum: IEnumVARIANT readonly dispid -4;
    property Item[Index: Integer]: IPCsText readonly dispid 0; default;
    property Count: Integer readonly dispid 203;
    function Add(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Value: WideString): IPCsText; dispid 204;
    function FindByValue(const Value: WideString): IPCsText; dispid 205;
    procedure Remove(Index: Integer); dispid 201;
    function AddXYZ(X: Integer; Y: Integer; Z: Integer; const Value: WideString): IPCsText; dispid 202;
  end;

// *********************************************************************//
// Interface: IPCsSymbolTexts
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {A6BD811D-303B-43D2-8942-EEDCC531EEBA}
// *********************************************************************//
  IPCsSymbolTexts = interface(IPCsTexts)
    ['{A6BD811D-303B-43D2-8942-EEDCC531EEBA}']
    function Get_NameText: IPCsNameText; safecall;
    function Get_ArticleText: IPCsText; safecall;
    function Get_TypeText: IPCsText; safecall;
    function Get_FunctionText: IPCsText; safecall;
    function FindByType(AType: PsTextType): IPCsText; safecall;
    property NameText: IPCsNameText read Get_NameText;
    property ArticleText: IPCsText read Get_ArticleText;
    property TypeText: IPCsText read Get_TypeText;
    property FunctionText: IPCsText read Get_FunctionText;
  end;

// *********************************************************************//
// DispIntf:  IPCsSymbolTextsDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {A6BD811D-303B-43D2-8942-EEDCC531EEBA}
// *********************************************************************//
  IPCsSymbolTextsDisp = dispinterface
    ['{A6BD811D-303B-43D2-8942-EEDCC531EEBA}']
    property NameText: IPCsNameText readonly dispid 301;
    property ArticleText: IPCsText readonly dispid 302;
    property TypeText: IPCsText readonly dispid 303;
    property FunctionText: IPCsText readonly dispid 304;
    function FindByType(AType: PsTextType): IPCsText; dispid 305;
    property _NewEnum: IEnumVARIANT readonly dispid -4;
    property Item[Index: Integer]: IPCsText readonly dispid 0; default;
    property Count: Integer readonly dispid 203;
    function Add(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Value: WideString): IPCsText; dispid 204;
    function FindByValue(const Value: WideString): IPCsText; dispid 205;
    procedure Remove(Index: Integer); dispid 201;
    function AddXYZ(X: Integer; Y: Integer; Z: Integer; const Value: WideString): IPCsText; dispid 202;
  end;

// *********************************************************************//
// Interface: IPCsText
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {17859DD2-080B-4256-8F37-CFD149D1E6B7}
// *********************************************************************//
  IPCsText = interface(IPCsStatusPoint)
    ['{17859DD2-080B-4256-8F37-CFD149D1E6B7}']
    function Get_Value: WideString; safecall;
    procedure Set_Value(const Value: WideString); safecall;
    function Get_Origin: PsTextOrigin; safecall;
    procedure Set_Origin(Value: PsTextOrigin); safecall;
    function Get_Font: IPCsTextFont; safecall;
    function Get_TextType: PsTextOrigin; safecall;
    function Get_FieldWidth: Integer; safecall;
    procedure Set_FieldWidth(Value: Integer); safecall;
    function Get_WrapText: WordBool; safecall;
    procedure Set_WrapText(Value: WordBool); safecall;
    function Get_LinkID: Integer; safecall;
    procedure Set_LinkID(Value: Integer); safecall;
    function Get_PositionExt: TPCsPoint3D; safecall;
    procedure Set_PositionExt(Value: TPCsPoint3D); safecall;
    function Get_Color: PsPenColor; safecall;
    procedure Set_Color(Value: PsPenColor); safecall;
    function Get_Layer: IPCsLayer; safecall;
    procedure Set_Layer(const Value: IPCsLayer); safecall;
    function Get_DisplayText: WideString; safecall;
    function Get_LineCount: Integer; safecall;
    function Get_Rotation: Integer; safecall;
    procedure Set_Rotation(Value: Integer); safecall;
    function Get_IsBlank: WordBool; safecall;
    function GetRect(AFlags: psTextRectFlags): TPCsRect; safecall;
    procedure Delete; safecall;
    function Get_ObjectType: PsPCsObjectType; safecall;
    function Get_NotTranslated: WordBool; safecall;
    procedure Set_NotTranslated(Value: WordBool); safecall;
    property Value: WideString read Get_Value write Set_Value;
    property Origin: PsTextOrigin read Get_Origin write Set_Origin;
    property Font: IPCsTextFont read Get_Font;
    property TextType: PsTextOrigin read Get_TextType;
    property FieldWidth: Integer read Get_FieldWidth write Set_FieldWidth;
    property WrapText: WordBool read Get_WrapText write Set_WrapText;
    property LinkID: Integer read Get_LinkID write Set_LinkID;
    property PositionExt: TPCsPoint3D read Get_PositionExt write Set_PositionExt;
    property Color: PsPenColor read Get_Color write Set_Color;
    property Layer: IPCsLayer read Get_Layer write Set_Layer;
    property DisplayText: WideString read Get_DisplayText;
    property LineCount: Integer read Get_LineCount;
    property Rotation: Integer read Get_Rotation write Set_Rotation;
    property IsBlank: WordBool read Get_IsBlank;
    property ObjectType: PsPCsObjectType read Get_ObjectType;
    property NotTranslated: WordBool read Get_NotTranslated write Set_NotTranslated;
  end;

// *********************************************************************//
// DispIntf:  IPCsTextDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {17859DD2-080B-4256-8F37-CFD149D1E6B7}
// *********************************************************************//
  IPCsTextDisp = dispinterface
    ['{17859DD2-080B-4256-8F37-CFD149D1E6B7}']
    property Value: WideString dispid 401;
    property Origin: PsTextOrigin dispid 402;
    property Font: IPCsTextFont readonly dispid 403;
    property TextType: PsTextOrigin readonly dispid 404;
    property FieldWidth: Integer dispid 405;
    property WrapText: WordBool dispid 406;
    property LinkID: Integer dispid 407;
    property PositionExt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant dispid 408;
    property Color: PsPenColor dispid 301;
    property Layer: IPCsLayer dispid 302;
    property DisplayText: WideString readonly dispid 303;
    property LineCount: Integer readonly dispid 304;
    property Rotation: Integer dispid 305;
    property IsBlank: WordBool readonly dispid 306;
    function GetRect(AFlags: psTextRectFlags): {NOT_OLEAUTO(TPCsRect)}OleVariant; dispid 307;
    procedure Delete; dispid 308;
    property ObjectType: PsPCsObjectType readonly dispid 450;
    property NotTranslated: WordBool dispid 309;
    property Visible: WordBool dispid 201;
    property Position: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant dispid 202;
    property LStatus: Integer dispid 203;
    function Move(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Figure: IPCsFigure): WordBool; dispid 204;
    procedure Rotate(DeltaAngle: Integer; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 205;
    procedure Scale(Factor: Single; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 208;
    property Handle: LongWord readonly dispid 206;
    procedure GetPosition(out X: Integer; out Y: Integer; out Z: Integer); dispid 207;
    procedure SetPosition(X: Integer; Y: Integer; Z: Integer); dispid 209;
    property State: Integer dispid 250;
    function ChangeDrawingOrder(AAction: psChangeDrawingOrder): WordBool; dispid 210;
  end;

// *********************************************************************//
// Interface: IPCsTextFont
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {506D1B55-9E93-4FE7-9DC8-BDAEA92264B9}
// *********************************************************************//
  IPCsTextFont = interface(IDispatch)
    ['{506D1B55-9E93-4FE7-9DC8-BDAEA92264B9}']
    function Get_Name: WideString; safecall;
    procedure Set_Name(const Value: WideString); safecall;
    function Get_IsPCschematicFont: WordBool; safecall;
    function Get_Bold: WordBool; safecall;
    procedure Set_Bold(Value: WordBool); safecall;
    function Get_Italic: WordBool; safecall;
    procedure Set_Italic(Value: WordBool); safecall;
    function Get_Height: Integer; safecall;
    procedure Set_Height(Value: Integer); safecall;
    function Get_Width: Integer; safecall;
    procedure Set_Width(Value: Integer); safecall;
    property Name: WideString read Get_Name write Set_Name;
    property IsPCschematicFont: WordBool read Get_IsPCschematicFont;
    property Bold: WordBool read Get_Bold write Set_Bold;
    property Italic: WordBool read Get_Italic write Set_Italic;
    property Height: Integer read Get_Height write Set_Height;
    property Width: Integer read Get_Width write Set_Width;
  end;

// *********************************************************************//
// DispIntf:  IPCsTextFontDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {506D1B55-9E93-4FE7-9DC8-BDAEA92264B9}
// *********************************************************************//
  IPCsTextFontDisp = dispinterface
    ['{506D1B55-9E93-4FE7-9DC8-BDAEA92264B9}']
    property Name: WideString dispid 201;
    property IsPCschematicFont: WordBool readonly dispid 202;
    property Bold: WordBool dispid 203;
    property Italic: WordBool dispid 204;
    property Height: Integer dispid 205;
    property Width: Integer dispid 206;
  end;

// *********************************************************************//
// Interface: IPCsNameText
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {EDE5A210-B27A-44AA-BBAF-1BE2702A372A}
// *********************************************************************//
  IPCsNameText = interface(IPCsText)
    ['{EDE5A210-B27A-44AA-BBAF-1BE2702A372A}']
    function Get_RefDesFunction: WideString; safecall;
    procedure Set_RefDesFunction(const Value: WideString); safecall;
    function Get_RefDesLocation: WideString; safecall;
    procedure Set_RefDesLocation(const Value: WideString); safecall;
    function Get_RefDesOptions: psSymbolRefDesOption; safecall;
    procedure Set_RefDesOptions(Value: psSymbolRefDesOption); safecall;
    property RefDesFunction: WideString read Get_RefDesFunction write Set_RefDesFunction;
    property RefDesLocation: WideString read Get_RefDesLocation write Set_RefDesLocation;
    property RefDesOptions: psSymbolRefDesOption read Get_RefDesOptions write Set_RefDesOptions;
  end;

// *********************************************************************//
// DispIntf:  IPCsNameTextDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {EDE5A210-B27A-44AA-BBAF-1BE2702A372A}
// *********************************************************************//
  IPCsNameTextDisp = dispinterface
    ['{EDE5A210-B27A-44AA-BBAF-1BE2702A372A}']
    property RefDesFunction: WideString dispid 409;
    property RefDesLocation: WideString dispid 410;
    property RefDesOptions: psSymbolRefDesOption dispid 411;
    property Value: WideString dispid 401;
    property Origin: PsTextOrigin dispid 402;
    property Font: IPCsTextFont readonly dispid 403;
    property TextType: PsTextOrigin readonly dispid 404;
    property FieldWidth: Integer dispid 405;
    property WrapText: WordBool dispid 406;
    property LinkID: Integer dispid 407;
    property PositionExt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant dispid 408;
    property Color: PsPenColor dispid 301;
    property Layer: IPCsLayer dispid 302;
    property DisplayText: WideString readonly dispid 303;
    property LineCount: Integer readonly dispid 304;
    property Rotation: Integer dispid 305;
    property IsBlank: WordBool readonly dispid 306;
    function GetRect(AFlags: psTextRectFlags): {NOT_OLEAUTO(TPCsRect)}OleVariant; dispid 307;
    procedure Delete; dispid 308;
    property ObjectType: PsPCsObjectType readonly dispid 450;
    property NotTranslated: WordBool dispid 309;
    property Visible: WordBool dispid 201;
    property Position: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant dispid 202;
    property LStatus: Integer dispid 203;
    function Move(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Figure: IPCsFigure): WordBool; dispid 204;
    procedure Rotate(DeltaAngle: Integer; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 205;
    procedure Scale(Factor: Single; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 208;
    property Handle: LongWord readonly dispid 206;
    procedure GetPosition(out X: Integer; out Y: Integer; out Z: Integer); dispid 207;
    procedure SetPosition(X: Integer; Y: Integer; Z: Integer); dispid 209;
    property State: Integer dispid 250;
    function ChangeDrawingOrder(AAction: psChangeDrawingOrder): WordBool; dispid 210;
  end;

// *********************************************************************//
// Interface: IPCsSymbol
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {9F12909F-4EE6-40C7-89C7-5B110A48355A}
// *********************************************************************//
  IPCsSymbol = interface(IPCsStatusPoint)
    ['{9F12909F-4EE6-40C7-89C7-5B110A48355A}']
    function Get_LibType: IPCsLibType; safecall;
    function Get_IsVMirrored: WordBool; safecall;
    procedure Set_IsVMirrored(Value: WordBool); safecall;
    function Get_IsHMirrored: WordBool; safecall;
    procedure Set_IsHMirrored(Value: WordBool); safecall;
    function Get_SymbolTexts: IPCsSymbolTexts; safecall;
    function Get_Page: IPCsPage; safecall;
    function Get_Layer: IPCsLayer; safecall;
    procedure Set_Layer(const Value: IPCsLayer); safecall;
    function Get_FullName: WideString; safecall;
    procedure Set_FullName(const Value: WideString); safecall;
    function GetRect(AType: PsSymbolRectType): TPCsRect; safecall;
    function Get_Connections: IPCsConnections; safecall;
    function Get_Rotation: Integer; safecall;
    procedure Set_Rotation(Value: Integer); safecall;
    function Get_ScaleFactor: Double; safecall;
    procedure Set_ScaleFactor(Value: Double); safecall;
    function Explode(const ADestination: IPCsFigure): WordBool; safecall;
    function Get_SymbolType1: psSymbolType; safecall;
    procedure Set_SymbolType1(Value: psSymbolType); safecall;
    function Get_SymbolType2: psSymbolType; safecall;
    procedure Set_SymbolType2(Value: psSymbolType); safecall;
    procedure ApplicationDataWriteString(const ASection: WideString; const AIdent: WideString;
                                         const AValue: WideString); safecall;
    function ApplicationDataReadString(const ASection: WideString; const AIdent: WideString;
                                       const ADefault: WideString): WideString; safecall;
    function Get_Datafields: IPCsDatafields; safecall;
    procedure Delete; safecall;
    function Get_ObjectType: PsPCsObjectType; safecall;
    function Get_ComponentGroupNo: Integer; safecall;
    procedure Set_ComponentGroupNo(Value: Integer); safecall;
    function Get_ComponentPosNo: Integer; safecall;
    procedure Set_ComponentPosNo(Value: Integer); safecall;
    function Get_Reference: IPCsReference; safecall;
    function Get_WithReference: WordBool; safecall;
    procedure Set_WithReference(Value: WordBool); safecall;
    function UpdateFromDatabase(Reserved: Integer): WordBool; safecall;
    property LibType: IPCsLibType read Get_LibType;
    property IsVMirrored: WordBool read Get_IsVMirrored write Set_IsVMirrored;
    property IsHMirrored: WordBool read Get_IsHMirrored write Set_IsHMirrored;
    property SymbolTexts: IPCsSymbolTexts read Get_SymbolTexts;
    property Page: IPCsPage read Get_Page;
    property Layer: IPCsLayer read Get_Layer write Set_Layer;
    property FullName: WideString read Get_FullName write Set_FullName;
    property Connections: IPCsConnections read Get_Connections;
    property Rotation: Integer read Get_Rotation write Set_Rotation;
    property ScaleFactor: Double read Get_ScaleFactor write Set_ScaleFactor;
    property SymbolType1: psSymbolType read Get_SymbolType1 write Set_SymbolType1;
    property SymbolType2: psSymbolType read Get_SymbolType2 write Set_SymbolType2;
    property Datafields: IPCsDatafields read Get_Datafields;
    property ObjectType: PsPCsObjectType read Get_ObjectType;
    property ComponentGroupNo: Integer read Get_ComponentGroupNo write Set_ComponentGroupNo;
    property ComponentPosNo: Integer read Get_ComponentPosNo write Set_ComponentPosNo;
    property Reference: IPCsReference read Get_Reference;
    property WithReference: WordBool read Get_WithReference write Set_WithReference;
  end;

// *********************************************************************//
// DispIntf:  IPCsSymbolDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {9F12909F-4EE6-40C7-89C7-5B110A48355A}
// *********************************************************************//
  IPCsSymbolDisp = dispinterface
    ['{9F12909F-4EE6-40C7-89C7-5B110A48355A}']
    property LibType: IPCsLibType readonly dispid 402;
    property IsVMirrored: WordBool dispid 405;
    property IsHMirrored: WordBool dispid 406;
    property SymbolTexts: IPCsSymbolTexts readonly dispid 407;
    property Page: IPCsPage readonly dispid 408;
    property Layer: IPCsLayer dispid 409;
    property FullName: WideString dispid 302;
    function GetRect(AType: PsSymbolRectType): {NOT_OLEAUTO(TPCsRect)}OleVariant; dispid 303;
    property Connections: IPCsConnections readonly dispid 304;
    property Rotation: Integer dispid 305;
    property ScaleFactor: Double dispid 306;
    function Explode(const ADestination: IPCsFigure): WordBool; dispid 307;
    property SymbolType1: psSymbolType dispid 301;
    property SymbolType2: psSymbolType dispid 308;
    procedure ApplicationDataWriteString(const ASection: WideString; const AIdent: WideString;
                                         const AValue: WideString); dispid 309;
    function ApplicationDataReadString(const ASection: WideString; const AIdent: WideString;
                                       const ADefault: WideString): WideString; dispid 310;
    property Datafields: IPCsDatafields readonly dispid 311;
    procedure Delete; dispid 312;
    property ObjectType: PsPCsObjectType readonly dispid 450;
    property ComponentGroupNo: Integer dispid 313;
    property ComponentPosNo: Integer dispid 314;
    property Reference: IPCsReference readonly dispid 315;
    property WithReference: WordBool dispid 316;
    function UpdateFromDatabase(Reserved: Integer): WordBool; dispid 317;
    property Visible: WordBool dispid 201;
    property Position: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant dispid 202;
    property LStatus: Integer dispid 203;
    function Move(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Figure: IPCsFigure): WordBool; dispid 204;
    procedure Rotate(DeltaAngle: Integer; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 205;
    procedure Scale(Factor: Single; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 208;
    property Handle: LongWord readonly dispid 206;
    procedure GetPosition(out X: Integer; out Y: Integer; out Z: Integer); dispid 207;
    procedure SetPosition(X: Integer; Y: Integer; Z: Integer); dispid 209;
    property State: Integer dispid 250;
    function ChangeDrawingOrder(AAction: psChangeDrawingOrder): WordBool; dispid 210;
  end;

// *********************************************************************//
// Interface: IPCsConnections
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {9CC2A147-2E7D-4613-B505-10B5AAD7AD2D}
// *********************************************************************//
  IPCsConnections = interface(IDispatch)
    ['{9CC2A147-2E7D-4613-B505-10B5AAD7AD2D}']
    function Get_Item(Index: Integer): IPCsConnection; safecall;
    function Get_Count: Integer; safecall;
    function Get__NewEnum: IEnumVARIANT; safecall;
    property Item[Index: Integer]: IPCsConnection read Get_Item; default;
    property Count: Integer read Get_Count;
    property _NewEnum: IEnumVARIANT read Get__NewEnum;
  end;

// *********************************************************************//
// DispIntf:  IPCsConnectionsDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {9CC2A147-2E7D-4613-B505-10B5AAD7AD2D}
// *********************************************************************//
  IPCsConnectionsDisp = dispinterface
    ['{9CC2A147-2E7D-4613-B505-10B5AAD7AD2D}']
    property Item[Index: Integer]: IPCsConnection readonly dispid 0; default;
    property Count: Integer readonly dispid 402;
    property _NewEnum: IEnumVARIANT readonly dispid -4;
  end;

// *********************************************************************//
// Interface: IPCsConnection
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {C5A27EFD-744C-4DCD-B4B5-1E4AC44023C7}
// *********************************************************************//
  IPCsConnection = interface(IPCsJoinPoint)
    ['{C5A27EFD-744C-4DCD-B4B5-1E4AC44023C7}']
    function Get_IsElectrical: WordBool; safecall;
    procedure Set_IsElectrical(Value: WordBool); safecall;
    function Get_IOStatus: TPCsIOStatus; safecall;
    procedure Set_IOStatus(Value: TPCsIOStatus); safecall;
    function Get_IsJumperingLink: WordBool; safecall;
    procedure Set_IsJumperingLink(Value: WordBool); safecall;
    function Get_IsCableScreen: WordBool; safecall;
    procedure Set_IsCableScreen(Value: WordBool); safecall;
    function Get_ConnectionTexts: IPCsConnectionTexts; safecall;
    function Get_Layer: IPCsLayer; safecall;
    procedure Set_Layer(const Value: IPCsLayer); safecall;
    function Get_PinName: WideString; safecall;
    procedure Set_PinName(const Value: WideString); safecall;
    procedure InitConnectionTexts(const Drawing: IPCsDrawing); safecall;
    procedure Delete; safecall;
    function Get_WithReference: WordBool; safecall;
    procedure Set_WithReference(Value: WordBool); safecall;
    function Get_Reference: IPCsReference; safecall;
    property IsElectrical: WordBool read Get_IsElectrical write Set_IsElectrical;
    property IOStatus: TPCsIOStatus read Get_IOStatus write Set_IOStatus;
    property IsJumperingLink: WordBool read Get_IsJumperingLink write Set_IsJumperingLink;
    property IsCableScreen: WordBool read Get_IsCableScreen write Set_IsCableScreen;
    property ConnectionTexts: IPCsConnectionTexts read Get_ConnectionTexts;
    property Layer: IPCsLayer read Get_Layer write Set_Layer;
    property PinName: WideString read Get_PinName write Set_PinName;
    property WithReference: WordBool read Get_WithReference write Set_WithReference;
    property Reference: IPCsReference read Get_Reference;
  end;

// *********************************************************************//
// DispIntf:  IPCsConnectionDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {C5A27EFD-744C-4DCD-B4B5-1E4AC44023C7}
// *********************************************************************//
  IPCsConnectionDisp = dispinterface
    ['{C5A27EFD-744C-4DCD-B4B5-1E4AC44023C7}']
    property IsElectrical: WordBool dispid 401;
    property IOStatus: {NOT_OLEAUTO(TPCsIOStatus)}OleVariant dispid 402;
    property IsJumperingLink: WordBool dispid 403;
    property IsCableScreen: WordBool dispid 404;
    property ConnectionTexts: IPCsConnectionTexts readonly dispid 405;
    property Layer: IPCsLayer dispid 406;
    property PinName: WideString dispid 407;
    procedure InitConnectionTexts(const Drawing: IPCsDrawing); dispid 408;
    procedure Delete; dispid 409;
    property WithReference: WordBool dispid 410;
    property Reference: IPCsReference readonly dispid 411;
    property Visible: WordBool dispid 201;
    property Position: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant dispid 202;
    property LStatus: Integer dispid 203;
    function Move(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Figure: IPCsFigure): WordBool; dispid 204;
    procedure Rotate(DeltaAngle: Integer; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 205;
    procedure Scale(Factor: Single; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 208;
    property Handle: LongWord readonly dispid 206;
    procedure GetPosition(out X: Integer; out Y: Integer; out Z: Integer); dispid 207;
    procedure SetPosition(X: Integer; Y: Integer; Z: Integer); dispid 209;
    property State: Integer dispid 250;
    function ChangeDrawingOrder(AAction: psChangeDrawingOrder): WordBool; dispid 210;
  end;

// *********************************************************************//
// Interface: IPCsConnectionTexts
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {9C483B8F-5BE6-4675-BDDA-8984012826B8}
// *********************************************************************//
  IPCsConnectionTexts = interface(IDispatch)
    ['{9C483B8F-5BE6-4675-BDDA-8984012826B8}']
    function Get_NameText: IPCsText; safecall;
    function Get_FunctionText: IPCsText; safecall;
    function Get_LabelText: IPCsText; safecall;
    function Get_DescriptionText: IPCsText; safecall;
    property NameText: IPCsText read Get_NameText;
    property FunctionText: IPCsText read Get_FunctionText;
    property LabelText: IPCsText read Get_LabelText;
    property DescriptionText: IPCsText read Get_DescriptionText;
  end;

// *********************************************************************//
// DispIntf:  IPCsConnectionTextsDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {9C483B8F-5BE6-4675-BDDA-8984012826B8}
// *********************************************************************//
  IPCsConnectionTextsDisp = dispinterface
    ['{9C483B8F-5BE6-4675-BDDA-8984012826B8}']
    property NameText: IPCsText readonly dispid 201;
    property FunctionText: IPCsText readonly dispid 202;
    property LabelText: IPCsText readonly dispid 203;
    property DescriptionText: IPCsText readonly dispid 204;
  end;

// *********************************************************************//
// Interface: IPCsDatafields
// Flags:     (320) Dual OleAutomation
// GUID:      {F4E0CF05-3AA2-4D46-B9AD-C199940CE630}
// *********************************************************************//
  IPCsDatafields = interface(IUnknown)
    ['{F4E0CF05-3AA2-4D46-B9AD-C199940CE630}']
    function Add(Pt: TPCsPoint3D; const Value: WideString): IPCsDatafield; safecall;
    function Get_Item(Index: Integer): IPCsDatafield; safecall;
    function Get_Count: Integer; safecall;
    function FindByValue(const Value: WideString): IPCsDatafield; safecall;
    function Get__NewEnum: IEnumVARIANT; safecall;
    property Item[Index: Integer]: IPCsDatafield read Get_Item; default;
    property Count: Integer read Get_Count;
    property _NewEnum: IEnumVARIANT read Get__NewEnum;
  end;

// *********************************************************************//
// DispIntf:  IPCsDatafieldsDisp
// Flags:     (320) Dual OleAutomation
// GUID:      {F4E0CF05-3AA2-4D46-B9AD-C199940CE630}
// *********************************************************************//
  IPCsDatafieldsDisp = dispinterface
    ['{F4E0CF05-3AA2-4D46-B9AD-C199940CE630}']
    function Add(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Value: WideString): IPCsDatafield; dispid 101;
    property Item[Index: Integer]: IPCsDatafield readonly dispid 0; default;
    property Count: Integer readonly dispid 102;
    function FindByValue(const Value: WideString): IPCsDatafield; dispid 103;
    property _NewEnum: IEnumVARIANT readonly dispid -4;
  end;

// *********************************************************************//
// Interface: IPCsDatafield
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {90B2DFB9-8544-4BFE-9ECD-3E23B69ECF7D}
// *********************************************************************//
  IPCsDatafield = interface(IPCsText)
    ['{90B2DFB9-8544-4BFE-9ECD-3E23B69ECF7D}']
    function Get_Pretext: WideString; safecall;
    procedure Set_Pretext(const Value: WideString); safecall;
    function Get_DataNo: Integer; safecall;
    procedure SetDataNo(DataNo: Integer; const DataLink: IPCsDataLink); safecall;
    function Get_FieldName: WideString; safecall;
    procedure Set_FieldName(const Value: WideString); safecall;
    property Pretext: WideString read Get_Pretext write Set_Pretext;
    property DataNo: Integer read Get_DataNo;
    property FieldName: WideString read Get_FieldName write Set_FieldName;
  end;

// *********************************************************************//
// DispIntf:  IPCsDatafieldDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {90B2DFB9-8544-4BFE-9ECD-3E23B69ECF7D}
// *********************************************************************//
  IPCsDatafieldDisp = dispinterface
    ['{90B2DFB9-8544-4BFE-9ECD-3E23B69ECF7D}']
    property Pretext: WideString dispid 409;
    property DataNo: Integer readonly dispid 410;
    procedure SetDataNo(DataNo: Integer; const DataLink: IPCsDataLink); dispid 411;
    property FieldName: WideString dispid 412;
    property Value: WideString dispid 401;
    property Origin: PsTextOrigin dispid 402;
    property Font: IPCsTextFont readonly dispid 403;
    property TextType: PsTextOrigin readonly dispid 404;
    property FieldWidth: Integer dispid 405;
    property WrapText: WordBool dispid 406;
    property LinkID: Integer dispid 407;
    property PositionExt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant dispid 408;
    property Color: PsPenColor dispid 301;
    property Layer: IPCsLayer dispid 302;
    property DisplayText: WideString readonly dispid 303;
    property LineCount: Integer readonly dispid 304;
    property Rotation: Integer dispid 305;
    property IsBlank: WordBool readonly dispid 306;
    function GetRect(AFlags: psTextRectFlags): {NOT_OLEAUTO(TPCsRect)}OleVariant; dispid 307;
    procedure Delete; dispid 308;
    property ObjectType: PsPCsObjectType readonly dispid 450;
    property NotTranslated: WordBool dispid 309;
    property Visible: WordBool dispid 201;
    property Position: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant dispid 202;
    property LStatus: Integer dispid 203;
    function Move(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Figure: IPCsFigure): WordBool; dispid 204;
    procedure Rotate(DeltaAngle: Integer; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 205;
    procedure Scale(Factor: Single; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 208;
    property Handle: LongWord readonly dispid 206;
    procedure GetPosition(out X: Integer; out Y: Integer; out Z: Integer); dispid 207;
    procedure SetPosition(X: Integer; Y: Integer; Z: Integer); dispid 209;
    property State: Integer dispid 250;
    function ChangeDrawingOrder(AAction: psChangeDrawingOrder): WordBool; dispid 210;
  end;

// *********************************************************************//
// Interface: IPCsDataLink
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {A07DB9EC-882F-4041-8B69-A16E8AD3BA8D}
// *********************************************************************//
  IPCsDataLink = interface(IPCsObject)
    ['{A07DB9EC-882F-4041-8B69-A16E8AD3BA8D}']
  end;

// *********************************************************************//
// DispIntf:  IPCsDataLinkDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {A07DB9EC-882F-4041-8B69-A16E8AD3BA8D}
// *********************************************************************//
  IPCsDataLinkDisp = dispinterface
    ['{A07DB9EC-882F-4041-8B69-A16E8AD3BA8D}']
    function HasSameOwnerDrawing(const AObject: IPCsObject): WordBool; dispid 190;
    property OwnerDrawing: IPCsDrawing readonly dispid 191;
  end;

// *********************************************************************//
// Interface: IPCsPageSetup
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {8462D121-9F59-4C75-B741-499FEF19488C}
// *********************************************************************//
  IPCsPageSetup = interface(IDispatch)
    ['{8462D121-9F59-4C75-B741-499FEF19488C}']
    function Get_Orientation: psOrientation; safecall;
    procedure Set_Orientation(Value: psOrientation); safecall;
    function Get_PaperSize: WideString; safecall;
    procedure Set_PaperSize(const Value: WideString); safecall;
    function Get_PaperWidth: Integer; safecall;
    procedure Set_PaperWidth(Value: Integer); safecall;
    function Get_PaperHeight: Integer; safecall;
    procedure Set_PaperHeight(Value: Integer); safecall;
    property Orientation: psOrientation read Get_Orientation write Set_Orientation;
    property PaperSize: WideString read Get_PaperSize write Set_PaperSize;
    property PaperWidth: Integer read Get_PaperWidth write Set_PaperWidth;
    property PaperHeight: Integer read Get_PaperHeight write Set_PaperHeight;
  end;

// *********************************************************************//
// DispIntf:  IPCsPageSetupDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {8462D121-9F59-4C75-B741-499FEF19488C}
// *********************************************************************//
  IPCsPageSetupDisp = dispinterface
    ['{8462D121-9F59-4C75-B741-499FEF19488C}']
    property Orientation: psOrientation dispid 201;
    property PaperSize: WideString dispid 202;
    property PaperWidth: Integer dispid 203;
    property PaperHeight: Integer dispid 204;
  end;

// *********************************************************************//
// Interface: IPCsPageReference
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {A845D055-E632-48E9-B3C6-BCA6C2772518}
// *********************************************************************//
  IPCsPageReference = interface(IDispatch)
    ['{A845D055-E632-48E9-B3C6-BCA6C2772518}']
    function Get_HorisontalReference: IPCsBasePageRef; safecall;
    function Get_VerticalReference: IPCsBasePageRef; safecall;
    function Get_IOMainReference: psDirection; safecall;
    procedure Set_IOMainReference(Value: psDirection); safecall;
    property HorisontalReference: IPCsBasePageRef read Get_HorisontalReference;
    property VerticalReference: IPCsBasePageRef read Get_VerticalReference;
    property IOMainReference: psDirection read Get_IOMainReference write Set_IOMainReference;
  end;

// *********************************************************************//
// DispIntf:  IPCsPageReferenceDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {A845D055-E632-48E9-B3C6-BCA6C2772518}
// *********************************************************************//
  IPCsPageReferenceDisp = dispinterface
    ['{A845D055-E632-48E9-B3C6-BCA6C2772518}']
    property HorisontalReference: IPCsBasePageRef readonly dispid 401;
    property VerticalReference: IPCsBasePageRef readonly dispid 402;
    property IOMainReference: psDirection dispid 403;
  end;

// *********************************************************************//
// Interface: IPCsBasePageRef
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {7C95FC24-0E20-4471-B5A7-DE41C4FC9772}
// *********************************************************************//
  IPCsBasePageRef = interface(IDispatch)
    ['{7C95FC24-0E20-4471-B5A7-DE41C4FC9772}']
    function Get_Active: WordBool; safecall;
    procedure Set_Active(Value: WordBool); safecall;
    function Get_StringFirstReference: WideString; safecall;
    procedure Set_StringFirstReference(const Value: WideString); safecall;
    function Get_ReferenceSymbolPosition: Integer; safecall;
    procedure Set_ReferenceSymbolPosition(Value: Integer); safecall;
    function Get_Position: TPCsPoint2D; safecall;
    procedure Set_Position(Value: TPCsPoint2D); safecall;
    property Active: WordBool read Get_Active write Set_Active;
    property StringFirstReference: WideString read Get_StringFirstReference write Set_StringFirstReference;
    property ReferenceSymbolPosition: Integer read Get_ReferenceSymbolPosition write Set_ReferenceSymbolPosition;
    property Position: TPCsPoint2D read Get_Position write Set_Position;
  end;

// *********************************************************************//
// DispIntf:  IPCsBasePageRefDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {7C95FC24-0E20-4471-B5A7-DE41C4FC9772}
// *********************************************************************//
  IPCsBasePageRefDisp = dispinterface
    ['{7C95FC24-0E20-4471-B5A7-DE41C4FC9772}']
    property Active: WordBool dispid 401;
    property StringFirstReference: WideString dispid 402;
    property ReferenceSymbolPosition: Integer dispid 403;
    property Position: {NOT_OLEAUTO(TPCsPoint2D)}OleVariant dispid 404;
  end;

// *********************************************************************//
// Interface: IPCsExternalObjects
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {AE4237D8-E4BE-427B-AC4E-14D3EB492067}
// *********************************************************************//
  IPCsExternalObjects = interface(IDispatch)
    ['{AE4237D8-E4BE-427B-AC4E-14D3EB492067}']
    procedure AddMetafile(Position: TPCsPoint3D; EnhMetafileHandle: Integer; Size: TPCsPoint2D); safecall;
  end;

// *********************************************************************//
// DispIntf:  IPCsExternalObjectsDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {AE4237D8-E4BE-427B-AC4E-14D3EB492067}
// *********************************************************************//
  IPCsExternalObjectsDisp = dispinterface
    ['{AE4237D8-E4BE-427B-AC4E-14D3EB492067}']
    procedure AddMetafile(Position: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant;
                          EnhMetafileHandle: Integer; Size: {NOT_OLEAUTO(TPCsPoint2D)}OleVariant); dispid 202;
  end;

// *********************************************************************//
// Interface: IPCsLayers
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {FEAF952B-B3F0-4671-AF8B-2026453A177E}
// *********************************************************************//
  IPCsLayers = interface(IDispatch)
    ['{FEAF952B-B3F0-4671-AF8B-2026453A177E}']
    function Get__NewEnum: IEnumVARIANT; safecall;
    function Get_Item(Index: Integer): IPCsLayer; safecall;
    function Get_Count: Integer; safecall;
    function FindByName(const Name: WideString): IPCsLayer; safecall;
    property _NewEnum: IEnumVARIANT read Get__NewEnum;
    property Item[Index: Integer]: IPCsLayer read Get_Item; default;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  IPCsLayersDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {FEAF952B-B3F0-4671-AF8B-2026453A177E}
// *********************************************************************//
  IPCsLayersDisp = dispinterface
    ['{FEAF952B-B3F0-4671-AF8B-2026453A177E}']
    property _NewEnum: IEnumVARIANT readonly dispid -4;
    property Item[Index: Integer]: IPCsLayer readonly dispid 0; default;
    property Count: Integer readonly dispid 203;
    function FindByName(const Name: WideString): IPCsLayer; dispid 204;
  end;

// *********************************************************************//
// Interface: IPCsApplication
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {226A846F-E175-4CE5-8C6F-B57EC56EEAB6}
// *********************************************************************//
  IPCsApplication = interface(IDispatch)
    ['{226A846F-E175-4CE5-8C6F-B57EC56EEAB6}']
    function Get_Documents: IPCsDocuments; safecall;
    function Get_AvailableFonts: OleVariant; safecall;
    function Get_PenColors(PenColor: PsPenColor): OLE_COLOR; safecall;
    function Get_Preferences: IPCsApplicationPreferences; safecall;
    function Get_SymbolLibraryMenu: IPCsSymbolMenu; safecall;
    function Get_StatusBar: WideString; safecall;
    procedure Set_StatusBar(const Value: WideString); safecall;
    function Get_Visible: WordBool; safecall;
    procedure Set_Visible(Value: WordBool); safecall;
    procedure Quit; safecall;
    procedure DontShutDownOnLastRelease; safecall;
    function CreateVirtualPage: IPCsPage; safecall;
    function Get_ActiveComponentDatabase: IPCsComponentDatabase; safecall;
    function ExecuteModule(const Filename: WideString; const ParameterString: WideString): WordBool; safecall;
    procedure Redraw; safecall;
    function Get_ActiveDocument: IPCsDocument; safecall;
    procedure Set_ActiveDocument(const Value: IPCsDocument); safecall;
    procedure ZoomAll; safecall;
    function LoadString(Identifier: Integer): WideString; safecall;
    function SetOleStatusbarText(PanelIndex: Integer; const Text: WideString; Width: Integer;
                                 Alignment: Shortint): TWindowHandle; safecall;
    function Get_Dialogs(Kind: psDialogKind): IPCsDialog; safecall;
    procedure Activate; safecall;
    function Get_ExplorerWindow: IPCsExplorerWindow; safecall;
    function CreateNewLibType: IPCsLibType; safecall;
    property Documents: IPCsDocuments read Get_Documents;
    property AvailableFonts: OleVariant read Get_AvailableFonts;
    property PenColors[PenColor: PsPenColor]: OLE_COLOR read Get_PenColors;
    property Preferences: IPCsApplicationPreferences read Get_Preferences;
    property SymbolLibraryMenu: IPCsSymbolMenu read Get_SymbolLibraryMenu;
    property StatusBar: WideString read Get_StatusBar write Set_StatusBar;
    property Visible: WordBool read Get_Visible write Set_Visible;
    property ActiveComponentDatabase: IPCsComponentDatabase read Get_ActiveComponentDatabase;
    property ActiveDocument: IPCsDocument read Get_ActiveDocument write Set_ActiveDocument;
    property Dialogs[Kind: psDialogKind]: IPCsDialog read Get_Dialogs;
    property ExplorerWindow: IPCsExplorerWindow read Get_ExplorerWindow;
  end;

// *********************************************************************//
// DispIntf:  IPCsApplicationDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {226A846F-E175-4CE5-8C6F-B57EC56EEAB6}
// *********************************************************************//
  IPCsApplicationDisp = dispinterface
    ['{226A846F-E175-4CE5-8C6F-B57EC56EEAB6}']
    property Documents: IPCsDocuments readonly dispid 400;
    property AvailableFonts: OleVariant readonly dispid 402;
    property PenColors[PenColor: PsPenColor]: OLE_COLOR readonly dispid 403;
    property Preferences: IPCsApplicationPreferences readonly dispid 404;
    property SymbolLibraryMenu: IPCsSymbolMenu readonly dispid 405;
    property StatusBar: WideString dispid 406;
    property Visible: WordBool dispid 408;
    procedure Quit; dispid 409;
    procedure DontShutDownOnLastRelease; dispid 410;
    function CreateVirtualPage: IPCsPage; dispid 411;
    property ActiveComponentDatabase: IPCsComponentDatabase readonly dispid 412;
    function ExecuteModule(const Filename: WideString; const ParameterString: WideString): WordBool; dispid 413;
    procedure Redraw; dispid 414;
    property ActiveDocument: IPCsDocument dispid 201;
    procedure ZoomAll; dispid 202;
    function LoadString(Identifier: Integer): WideString; dispid 203;
    function SetOleStatusbarText(PanelIndex: Integer; const Text: WideString; Width: Integer;
                                 Alignment: Shortint): TWindowHandle; dispid 204;
    property Dialogs[Kind: psDialogKind]: IPCsDialog readonly dispid 205;
    procedure Activate; dispid 206;
    property ExplorerWindow: IPCsExplorerWindow readonly dispid 207;
    function CreateNewLibType: IPCsLibType; dispid 208;
  end;

// *********************************************************************//
// Interface: IPCsDocuments
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {5731AB31-7CA2-45F2-B84C-108E708E7FD0}
// *********************************************************************//
  IPCsDocuments = interface(IDispatch)
    ['{5731AB31-7CA2-45F2-B84C-108E708E7FD0}']
    function Add(const Template: WideString; const NewFilename: WideString): IPCsDocument; safecall;
    function Get_Item(Index: Integer): IPCsDocument; safecall;
    function Get_Count: Integer; safecall;
    function Get__NewEnum: IEnumVARIANT; safecall;
    function Open(const Filename: WideString): IPCsDocument; safecall;
    property Item[Index: Integer]: IPCsDocument read Get_Item; default;
    property Count: Integer read Get_Count;
    property _NewEnum: IEnumVARIANT read Get__NewEnum;
  end;

// *********************************************************************//
// DispIntf:  IPCsDocumentsDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {5731AB31-7CA2-45F2-B84C-108E708E7FD0}
// *********************************************************************//
  IPCsDocumentsDisp = dispinterface
    ['{5731AB31-7CA2-45F2-B84C-108E708E7FD0}']
    function Add(const Template: WideString; const NewFilename: WideString): IPCsDocument; dispid 201;
    property Item[Index: Integer]: IPCsDocument readonly dispid 0; default;
    property Count: Integer readonly dispid 203;
    property _NewEnum: IEnumVARIANT readonly dispid -4;
    function Open(const Filename: WideString): IPCsDocument; dispid 202;
  end;

// *********************************************************************//
// Interface: IPCsDocument
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {D7D52290-DB58-45FA-85BB-25FB073FA5AA}
// *********************************************************************//
  IPCsDocument = interface(IDispatch)
    ['{D7D52290-DB58-45FA-85BB-25FB073FA5AA}']
    function Get_Drawing: IPCsDrawing; safecall;
    function Get_Application: IPCsApplication; safecall;
    procedure Close(SaveChanges: WordBool; const Filename: WideString); safecall;
    function Get_FullName: WideString; safecall;
    function ChangeFilename(const Filename: WideString): WordBool; safecall;
    function Get_CharSetName: WideString; safecall;
    procedure Set_CharSetName(const Value: WideString); safecall;
    procedure Activate; safecall;
    function Get_WindowHandle: TWindowHandle; safecall;
    function Get_IsNewFile: WordBool; safecall;
    procedure ZoomAll; safecall;
    function Get_ActivePage: IPCsPage; safecall;
    procedure Set_ActivePage(const Value: IPCsPage); safecall;
    function Get_Saved: WordBool; safecall;
    procedure Set_Saved(Value: WordBool); safecall;
    function Get_DataLink: IPCsDataLink; safecall;
    procedure Save; safecall;
    procedure SaveAs(const Filename: WideString); safecall;
    procedure Print; safecall;
    function Get_SelectStatus: IPCsSelectStatus; safecall;
    function Get_Handle: LongWord; safecall;
    procedure BeginUndo(const Title: WideString); safecall;
    procedure EndUndo; safecall;
    procedure Undo; safecall;
    procedure Select(const Page: IPCsPage; const Layer: IPCsLayer; const FilterString: WideString); safecall;
    procedure SelectInWindow(AreaLo: TPCsPoint3D; AreaHi: TPCsPoint3D; const Page: IPCsPage;
                             const Layer: IPCsLayer; const FilterString: WideString); safecall;
    function GotoObject(const AObject: IUnknown; Options: psGotoObjectOption): WordBool; safecall;
    property Drawing: IPCsDrawing read Get_Drawing;
    property Application: IPCsApplication read Get_Application;
    property FullName: WideString read Get_FullName;
    property CharSetName: WideString read Get_CharSetName write Set_CharSetName;
    property WindowHandle: TWindowHandle read Get_WindowHandle;
    property IsNewFile: WordBool read Get_IsNewFile;
    property ActivePage: IPCsPage read Get_ActivePage write Set_ActivePage;
    property Saved: WordBool read Get_Saved write Set_Saved;
    property DataLink: IPCsDataLink read Get_DataLink;
    property SelectStatus: IPCsSelectStatus read Get_SelectStatus;
    property Handle: LongWord read Get_Handle;
  end;

// *********************************************************************//
// DispIntf:  IPCsDocumentDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {D7D52290-DB58-45FA-85BB-25FB073FA5AA}
// *********************************************************************//
  IPCsDocumentDisp = dispinterface
    ['{D7D52290-DB58-45FA-85BB-25FB073FA5AA}']
    property Drawing: IPCsDrawing readonly dispid 201;
    property Application: IPCsApplication readonly dispid 202;
    procedure Close(SaveChanges: WordBool; const Filename: WideString); dispid 203;
    property FullName: WideString readonly dispid 206;
    function ChangeFilename(const Filename: WideString): WordBool; dispid 207;
    property CharSetName: WideString dispid 208;
    procedure Activate; dispid 210;
    property WindowHandle: TWindowHandle readonly dispid 204;
    property IsNewFile: WordBool readonly dispid 211;
    procedure ZoomAll; dispid 212;
    property ActivePage: IPCsPage dispid 205;
    property Saved: WordBool dispid 209;
    property DataLink: IPCsDataLink readonly dispid 213;
    procedure Save; dispid 214;
    procedure SaveAs(const Filename: WideString); dispid 215;
    procedure Print; dispid 216;
    property SelectStatus: IPCsSelectStatus readonly dispid 217;
    property Handle: LongWord readonly dispid 218;
    procedure BeginUndo(const Title: WideString); dispid 219;
    procedure EndUndo; dispid 220;
    procedure Undo; dispid 221;
    procedure Select(const Page: IPCsPage; const Layer: IPCsLayer; const FilterString: WideString); dispid 222;
    procedure SelectInWindow(AreaLo: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant;
                             AreaHi: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Page: IPCsPage;
                             const Layer: IPCsLayer; const FilterString: WideString); dispid 223;
    function GotoObject(const AObject: IUnknown; Options: psGotoObjectOption): WordBool; dispid 224;
  end;

// *********************************************************************//
// Interface: IPCsApplicationPreferences
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0042104A-A21F-4DFB-A60A-A078BA62E955}
// *********************************************************************//
  IPCsApplicationPreferences = interface(IDispatch)
    ['{0042104A-A21F-4DFB-A60A-A078BA62E955}']
    function Get_Folders(FolderType: PsFolderType): WideString; safecall;
    procedure Set_Folders(FolderType: PsFolderType; const Value: WideString); safecall;
    function Get_DxfDwgMapFile: WideString; safecall;
    procedure Set_DxfDwgMapFile(const Value: WideString); safecall;
    function Get_SymbolFolderAliases: IPCsSymbolFolderAliases; safecall;
    property Folders[FolderType: PsFolderType]: WideString read Get_Folders write Set_Folders;
    property DxfDwgMapFile: WideString read Get_DxfDwgMapFile write Set_DxfDwgMapFile;
    property SymbolFolderAliases: IPCsSymbolFolderAliases read Get_SymbolFolderAliases;
  end;

// *********************************************************************//
// DispIntf:  IPCsApplicationPreferencesDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0042104A-A21F-4DFB-A60A-A078BA62E955}
// *********************************************************************//
  IPCsApplicationPreferencesDisp = dispinterface
    ['{0042104A-A21F-4DFB-A60A-A078BA62E955}']
    property Folders[FolderType: PsFolderType]: WideString dispid 201;
    property DxfDwgMapFile: WideString dispid 202;
    property SymbolFolderAliases: IPCsSymbolFolderAliases readonly dispid 203;
  end;

// *********************************************************************//
// Interface: IPCsApplicationPreferences2
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {106E13A7-E58E-4A3A-A257-F88C3029F1FD}
// *********************************************************************//
  IPCsApplicationPreferences2 = interface(IPCsApplicationPreferences)
    ['{106E13A7-E58E-4A3A-A257-F88C3029F1FD}']
    function Get_Database: IPCsDatabaseSetup; safecall;
    property Database: IPCsDatabaseSetup read Get_Database;
  end;

// *********************************************************************//
// DispIntf:  IPCsApplicationPreferences2Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {106E13A7-E58E-4A3A-A257-F88C3029F1FD}
// *********************************************************************//
  IPCsApplicationPreferences2Disp = dispinterface
    ['{106E13A7-E58E-4A3A-A257-F88C3029F1FD}']
    property Database: IPCsDatabaseSetup readonly dispid 301;
    property Folders[FolderType: PsFolderType]: WideString dispid 201;
    property DxfDwgMapFile: WideString dispid 202;
    property SymbolFolderAliases: IPCsSymbolFolderAliases readonly dispid 203;
  end;

// *********************************************************************//
// Interface: IPCsSymbolMenu
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {5ED0A430-81EC-4529-9D9B-A35EF0A3177E}
// *********************************************************************//
  IPCsSymbolMenu = interface(IDispatch)
    ['{5ED0A430-81EC-4529-9D9B-A35EF0A3177E}']
    function Get_InitialDir: WideString; safecall;
    procedure Set_InitialDir(const Value: WideString); safecall;
    function Get_Filename: WideString; safecall;
    procedure Set_Filename(const Value: WideString); safecall;
    function Get_DatabaseLoad: WordBool; safecall;
    procedure Set_DatabaseLoad(Value: WordBool); safecall;
    function Execute: Integer; safecall;
    property InitialDir: WideString read Get_InitialDir write Set_InitialDir;
    property Filename: WideString read Get_Filename write Set_Filename;
    property DatabaseLoad: WordBool read Get_DatabaseLoad write Set_DatabaseLoad;
  end;

// *********************************************************************//
// DispIntf:  IPCsSymbolMenuDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {5ED0A430-81EC-4529-9D9B-A35EF0A3177E}
// *********************************************************************//
  IPCsSymbolMenuDisp = dispinterface
    ['{5ED0A430-81EC-4529-9D9B-A35EF0A3177E}']
    property InitialDir: WideString dispid 201;
    property Filename: WideString dispid 202;
    property DatabaseLoad: WordBool dispid 203;
    function Execute: Integer; dispid 204;
  end;

// *********************************************************************//
// Interface: IPCsComponentDatabase
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {C8B26975-AC5E-473F-B922-DA08C50C383E}
// *********************************************************************//
  IPCsComponentDatabase = interface(IDispatch)
    ['{C8B26975-AC5E-473F-B922-DA08C50C383E}']
    function Get_Name: Integer; safecall;
    function Get_Active: WordBool; safecall;
    function Get_TableComponents: WideString; safecall;
    function Get_DatabaseAlias: WideString; safecall;
    function Get_DatabaseFile: WideString; safecall;
    function Get_DatabaseString: WideString; safecall;
    property Name: Integer read Get_Name;
    property Active: WordBool read Get_Active;
    property TableComponents: WideString read Get_TableComponents;
    property DatabaseAlias: WideString read Get_DatabaseAlias;
    property DatabaseFile: WideString read Get_DatabaseFile;
    property DatabaseString: WideString read Get_DatabaseString;
  end;

// *********************************************************************//
// DispIntf:  IPCsComponentDatabaseDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {C8B26975-AC5E-473F-B922-DA08C50C383E}
// *********************************************************************//
  IPCsComponentDatabaseDisp = dispinterface
    ['{C8B26975-AC5E-473F-B922-DA08C50C383E}']
    property Name: Integer readonly dispid 201;
    property Active: WordBool readonly dispid 202;
    property TableComponents: WideString readonly dispid 203;
    property DatabaseAlias: WideString readonly dispid 204;
    property DatabaseFile: WideString readonly dispid 205;
    property DatabaseString: WideString readonly dispid 206;
  end;

// *********************************************************************//
// Interface: IPCsLibTypeConnections
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0573CF73-2CD7-4B66-BABA-4421EC53A302}
// *********************************************************************//
  IPCsLibTypeConnections = interface(IPCsConnections)
    ['{0573CF73-2CD7-4B66-BABA-4421EC53A302}']
    function Add(Pt: TPCsPoint3D): IPCsConnection; safecall;
  end;

// *********************************************************************//
// DispIntf:  IPCsLibTypeConnectionsDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0573CF73-2CD7-4B66-BABA-4421EC53A302}
// *********************************************************************//
  IPCsLibTypeConnectionsDisp = dispinterface
    ['{0573CF73-2CD7-4B66-BABA-4421EC53A302}']
    function Add(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant): IPCsConnection; dispid 501;
    property Item[Index: Integer]: IPCsConnection readonly dispid 0; default;
    property Count: Integer readonly dispid 402;
    property _NewEnum: IEnumVARIANT readonly dispid -4;
  end;

// *********************************************************************//
// Interface: IPCsArcs
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {64AF72DA-7A46-4033-BFE9-034BEEC408FF}
// *********************************************************************//
  IPCsArcs = interface(IDispatch)
    ['{64AF72DA-7A46-4033-BFE9-034BEEC408FF}']
    function Get_Item(Index: Integer): IPCsArc; safecall;
    function Get_Count: Integer; safecall;
    function Add(Center: TPCsPoint3D; Radius: Integer; StartAng: Integer; EndAng: Integer): IPCsArc; safecall;
    function Get__NewEnum: IEnumVARIANT; safecall;
    function AddXYZ(CenterX: Integer; CenterY: Integer; CenterZ: Integer; Radius: Integer;
                    StartAng: Integer; EndAng: Integer): IPCsArc; safecall;
    property Item[Index: Integer]: IPCsArc read Get_Item; default;
    property Count: Integer read Get_Count;
    property _NewEnum: IEnumVARIANT read Get__NewEnum;
  end;

// *********************************************************************//
// DispIntf:  IPCsArcsDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {64AF72DA-7A46-4033-BFE9-034BEEC408FF}
// *********************************************************************//
  IPCsArcsDisp = dispinterface
    ['{64AF72DA-7A46-4033-BFE9-034BEEC408FF}']
    property Item[Index: Integer]: IPCsArc readonly dispid 0; default;
    property Count: Integer readonly dispid 402;
    function Add(Center: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; Radius: Integer; StartAng: Integer;
                 EndAng: Integer): IPCsArc; dispid 403;
    property _NewEnum: IEnumVARIANT readonly dispid -4;
    function AddXYZ(CenterX: Integer; CenterY: Integer; CenterZ: Integer; Radius: Integer;
                    StartAng: Integer; EndAng: Integer): IPCsArc; dispid 404;
  end;

// *********************************************************************//
// Interface: IPCsArc
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {BC2F1F85-76B3-408D-860D-215F963B9A40}
// *********************************************************************//
  IPCsArc = interface(IPCsStatusPoint)
    ['{BC2F1F85-76B3-408D-860D-215F963B9A40}']
    function Get_IsFilled: WordBool; safecall;
    procedure Set_IsFilled(Value: WordBool); safecall;
    function Get_Radius: Integer; safecall;
    procedure Set_Radius(Value: Integer); safecall;
    function Get_StartAngle: Integer; safecall;
    procedure Set_StartAngle(Value: Integer); safecall;
    function Get_EndAngle: Integer; safecall;
    procedure Set_EndAngle(Value: Integer); safecall;
    function Get_PenWidth: Integer; safecall;
    procedure Set_PenWidth(Value: Integer); safecall;
    function Get__EllipseFactor: Integer; safecall;
    procedure Set__EllipseFactor(Value: Integer); safecall;
    function Get_Color: PsPenColor; safecall;
    procedure Set_Color(Value: PsPenColor); safecall;
    function Get_Layer: IPCsLayer; safecall;
    procedure Set_Layer(const Value: IPCsLayer); safecall;
    function Get_IsCircle: WordBool; safecall;
    function Get_Rotation: Integer; safecall;
    procedure Set_Rotation(Value: Integer); safecall;
    function Get_EllipseFactor: Double; safecall;
    procedure Set_EllipseFactor(Value: Double); safecall;
    procedure Delete; safecall;
    property IsFilled: WordBool read Get_IsFilled write Set_IsFilled;
    property Radius: Integer read Get_Radius write Set_Radius;
    property StartAngle: Integer read Get_StartAngle write Set_StartAngle;
    property EndAngle: Integer read Get_EndAngle write Set_EndAngle;
    property PenWidth: Integer read Get_PenWidth write Set_PenWidth;
    property _EllipseFactor: Integer read Get__EllipseFactor write Set__EllipseFactor;
    property Color: PsPenColor read Get_Color write Set_Color;
    property Layer: IPCsLayer read Get_Layer write Set_Layer;
    property IsCircle: WordBool read Get_IsCircle;
    property Rotation: Integer read Get_Rotation write Set_Rotation;
    property EllipseFactor: Double read Get_EllipseFactor write Set_EllipseFactor;
  end;

// *********************************************************************//
// DispIntf:  IPCsArcDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {BC2F1F85-76B3-408D-860D-215F963B9A40}
// *********************************************************************//
  IPCsArcDisp = dispinterface
    ['{BC2F1F85-76B3-408D-860D-215F963B9A40}']
    property IsFilled: WordBool dispid 401;
    property Radius: Integer dispid 403;
    property StartAngle: Integer dispid 404;
    property EndAngle: Integer dispid 405;
    property PenWidth: Integer dispid 407;
    property _EllipseFactor: Integer dispid 408;
    property Color: PsPenColor dispid 409;
    property Layer: IPCsLayer dispid 410;
    property IsCircle: WordBool readonly dispid 411;
    property Rotation: Integer dispid 412;
    property EllipseFactor: Double dispid 413;
    procedure Delete; dispid 414;
    property Visible: WordBool dispid 201;
    property Position: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant dispid 202;
    property LStatus: Integer dispid 203;
    function Move(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Figure: IPCsFigure): WordBool; dispid 204;
    procedure Rotate(DeltaAngle: Integer; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 205;
    procedure Scale(Factor: Single; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 208;
    property Handle: LongWord readonly dispid 206;
    procedure GetPosition(out X: Integer; out Y: Integer; out Z: Integer); dispid 207;
    procedure SetPosition(X: Integer; Y: Integer; Z: Integer); dispid 209;
    property State: Integer dispid 250;
    function ChangeDrawingOrder(AAction: psChangeDrawingOrder): WordBool; dispid 210;
  end;

// *********************************************************************//
// Interface: IApp
// Flags:     (4432) Hidden Dual OleAutomation Dispatchable
// GUID:      {C5E8FDA5-C88D-11D1-AE4B-0020AF16D64A}
// *********************************************************************//
  IApp = interface(IDispatch)
    ['{C5E8FDA5-C88D-11D1-AE4B-0020AF16D64A}']
    function CreateProject(const AProjectTitle: WideString; const AFilename: WideString): Integer; safecall;
    function OpenProject(const AFilename: WideString): Integer; safecall;
    procedure CloseProject(ProjectNo: Integer; SavedCheck: WordBool); safecall;
    procedure SaveWorkFile(ProjectNo: Integer; const AFilename: WideString); safecall;
    procedure GetProjectData(out Defs: OleVariant; out Data: OleVariant); safecall;
    procedure SetProjectData(Data: OleVariant); safecall;
    function OpenSymbol(const AFilename: WideString): Integer; safecall;
    procedure UpdateList(ListType: Integer); safecall;
    procedure CreateNetList(ListType: Integer; Layer: Integer); safecall;
    procedure FreeNetList; safecall;
    function GetNetListCount: Integer; safecall;
    procedure GetNetList(NetIndex: Integer; out AList: OleVariant); safecall;
    procedure CreateDynCableList(Layer: Integer); safecall;
    procedure FreeDynCableList; safecall;
    function GetDynCableListCount: Integer; safecall;
    procedure GetDynCableList(ListIndex: Integer; out AList: OleVariant); safecall;
    procedure NewPage(PageNo: Integer); safecall;
    procedure SetPageFunction(PageNo: Integer; AFunction: Integer); safecall;
    procedure GetPageBox(PageNo: Integer; out X1: Integer; out Y1: Integer; out X2: Integer;
                         out Y2: Integer); safecall;
    function GetPageCount: Integer; safecall;
    function PlaceSymbol(const SymbolFilename: WideString; Page: Integer; X: Integer; Y: Integer;
                         Z: Integer; Layer: Integer): Integer; safecall;
    function GetSymbolCount(PageNo: Integer): Integer; safecall;
    function GetSymbolHandle(PageNo: Integer; SymbolIndex: Integer): Integer; safecall;
    procedure RemoveSymbol(PageNo: Integer; SymbolHandle: Integer); safecall;
    procedure SetSymbolText(PageNo: Integer; SymbolHandle: Integer; const Name: WideString;
                            const TypeText: WideString; const Article: WideString;
                            const FunctionText: WideString); safecall;
    procedure SetSymbolType(PageNo: Integer; SymbolHandle: Integer; AType: Integer); safecall;
    function GetSymbolType(PageNo: Integer; SymbolHandle: Integer; out AType: Integer): WordBool; safecall;
    procedure GetSymbolText(PageNo: Integer; SymbolHandle: Integer; out Name: WideString;
                            out TypeText: WideString; out Article: WideString;
                            out FunctionText: WideString); safecall;
    procedure GetSymbolBox(PageNo: Integer; SymbolHandle: Integer; BoxType: Integer;
                           out X1: Integer; out Y1: Integer; out X2: Integer; out Y2: Integer); safecall;
    procedure MoveSymbol(PageNo: Integer; SymbolHandle: Integer; NewPage: Integer; X: Integer;
                         Y: Integer; Z: Integer; Layer: Integer); safecall;
    procedure GetSymbolPos(PageNo: Integer; SymbolHandle: Integer; out X: Integer; out Y: Integer;
                           out Z: Integer; out Layer: Integer); safecall;
    procedure GetSymbolConnectionPosRel(PageNo: Integer; SymbolHandle: Integer;
                                        ConnectionIndex: Integer; out X: Integer; out Y: Integer;
                                        out Z: Integer); safecall;
    procedure GetSymbolConnectionIOstatus(PageNo: Integer; SymbolHandle: Integer;
                                          ConnectionIndex: Integer; out AStatus: Integer); safecall;
    function PlaceLine(PageNo: Integer; Layer: Integer; StartX: Integer; StartY: Integer;
                       StartZ: Integer; EndX: Integer; EndY: Integer; EndZ: Integer): Integer; safecall;
    function PlacePolyLine(PageNo: Integer; Layer: Integer; AList: OleVariant): Integer; safecall;
    function GetLineCount(PageNo: Integer): Integer; safecall;
    function GetLineHandle(PageNo: Integer; LineIndex: Integer): Integer; safecall;
    procedure SetLineOptions(Color: Integer; Style: Integer; Width: Integer); safecall;
    procedure UpdateDrawing; safecall;
    function GetConnections(PageIndex: Integer; SymbolHandle: Integer): OleVariant; safecall;
    function PlaceDynamicCable(PageNo: Integer; Layer: Integer; Direction: Integer;
                               AList: OleVariant): Integer; safecall;
    function GetSymbolConnectionCount(PageNo: Integer; SymbolHandle: Integer): Integer; safecall;
    procedure CreateJoins(PageNo: Integer); safecall;
    function GetSymbolDataFieldCount(PageNo: Integer; SymbolHandle: Integer): Integer; safecall;
    procedure GetSymbolDataField(PageNo: Integer; SymbolHandle: Integer; FieldIndex: Integer;
                                 out AName: WideString; out AText: WideString); safecall;
    procedure SetSymbolDataField(PageNo: Integer; SymbolHandle: Integer; FieldIndex: Integer;
                                 const AText: WideString); safecall;
    procedure SetSymbolConnectionText(PageNo: Integer; SymbolHandle: Integer;
                                      ConnectionIndex: Integer; const Name: WideString;
                                      const FunctionText: WideString; const LabelText: WideString;
                                      const Description: WideString); safecall;
    procedure GetSymbolConnectionText(PageNo: Integer; SymbolHandle: Integer;
                                      ConnectionIndex: Integer; out Name: WideString;
                                      out FunctionText: WideString; out LabelText: WideString;
                                      out Description: WideString); safecall;
    function PlaceHeader(PageIndex: Integer; const HeaderFilename: WideString): WordBool; safecall;
    function PlaceDynamicSymbol(const SymbolFilename: WideString; PageNo: Integer; X: Integer;
                                Y: Integer; Z: Integer; Layer: Integer; AList: OleVariant): Integer; safecall;
    procedure GetSymbolFileName(PageNo: Integer; SymbolHandle: Integer; out AFilename: WideString); safecall;
    procedure GetSymbolTitle(PageNo: Integer; SymbolHandle: Integer; out ATitle: WideString); safecall;
    procedure SetSymbolNameData(PageNo: Integer; SymbolHandle: Integer; AWidth: Integer;
                                AHeight: Integer; AOrigin: Integer; AOutLine: Integer;
                                AColor: Integer; ABold: Integer; AVisible: Integer; AAngle: Integer); safecall;
    procedure SetSymbolNamePos(PageNo: Integer; SymbolHandle: Integer; X: Integer; Y: Integer;
                               Z: Integer); safecall;
    procedure SetSymbolTypeTextData(PageNo: Integer; SymbolHandle: Integer; AWidth: Integer;
                                    AHeight: Integer; AOrigin: Integer; AOutLine: Integer;
                                    AColor: Integer; ABold: Integer; AVisible: Integer;
                                    AAngle: Integer); safecall;
    procedure SetSymbolTypeTextPos(PageNo: Integer; SymbolHandle: Integer; X: Integer; Y: Integer;
                                   Z: Integer); safecall;
    procedure SetSymbolArticleData(PageNo: Integer; SymbolHandle: Integer; AWidth: Integer;
                                   AHeight: Integer; AOrigin: Integer; AOutLine: Integer;
                                   AColor: Integer; ABold: Integer; AVisible: Integer;
                                   AAngle: Integer); safecall;
    procedure SetSymbolArticlePos(PageNo: Integer; SymbolHandle: Integer; X: Integer; Y: Integer;
                                  Z: Integer); safecall;
    procedure SetSymbolFunctionData(PageNo: Integer; SymbolHandle: Integer; AWidth: Integer;
                                    AHeight: Integer; AOrigin: Integer; AOutLine: Integer;
                                    AColor: Integer; ABold: Integer; AVisible: Integer;
                                    AAngle: Integer); safecall;
    procedure SetSymbolFunctionPos(PageNo: Integer; SymbolHandle: Integer; X: Integer; Y: Integer;
                                   Z: Integer); safecall;
    procedure GetPageHeaders(out Headers: OleVariant); safecall;
    procedure GetSymbolFileNames(out FileNames: OleVariant); safecall;
    procedure GetTextWidth(ACharWidth: Integer; const AText: WideString; out ATextWidth: Integer); safecall;
    function AddSpeedButton(const ABarName: WideString; const AHint: WideString;
                            const Caption_BMP: WideString; WinHandle: Integer; CmdID: Integer): Integer; safecall;
    function AddMenuItem(const AOwnerMenu: WideString; const ACaption: WideString;
                         WinHandle: Integer; CmdID: Integer): Integer; safecall;
    function RemoveCmdItem(AItem: Integer): WordBool; safecall;
    function EnableCmdItem(AItem: Integer; AVisible: WordBool; AEnabled: WordBool): WordBool; safecall;
    procedure CreateTermList; safecall;
    procedure FreeTermList; safecall;
    function GetTermListCount: Integer; safecall;
    procedure GetTermList(ListIndex: Integer; out AList: OleVariant); safecall;
    procedure BringToFront; safecall;
    function GetProgramState: Integer; safecall;
    function AddMenu(const AOwnerMenu: WideString; const ACaption: WideString;
                     out AMenuName: WideString): Integer; safecall;
    procedure SetPageReference(PageNo: Integer; AFlags: Integer); safecall;
    procedure SetSymbolConnTextsVisible(PageNo: Integer; SymbolHandle: Integer; ATexts: Integer;
                                        AVisible: Integer); safecall;
    function PlaceText(const AText: WideString; Page: Integer; X: Integer; Y: Integer; Z: Integer;
                       Layer: Integer): Integer; safecall;
    procedure SetLayVisible(Layer: Integer; AVisible: Integer); safecall;
    procedure SymbolMenu(const ALibrary: WideString); safecall;
    procedure DontShutDownOnRelease; safecall;
    function GetProgramVersion: WideString; safecall;
    function GetProgramEnv(const EnvField: WideString): WideString; safecall;
    procedure UpdateMenus; safecall;
    procedure RotateSymbol(PageNo: Integer; SymbolHandle: Integer; Angle: Integer); safecall;
    procedure SetProjectTextsVisible(ATexts: Integer; AVisible: Integer); safecall;
    procedure TransferLibType(SrcProjectNo: Integer; DstProjectNo: Integer;
                              const AFilename: WideString); safecall;
    function GetActiveProjectHandle: Integer; safecall;
    function GetFunctionID(const FunctionName: WideString): Integer; safecall;
    procedure FunctionActivate(AID: Integer; Activate: WordBool); safecall;
    procedure FunctionVisible(AID: Integer; Visible: WordBool); safecall;
    procedure FunctionEnable(AID: Integer; Enable: WordBool); safecall;
    procedure SetFunctionMsg(AID: Integer; WinHandle: Integer; Msg: Integer); safecall;
    procedure GetProjectName(out AFilename: WideString); safecall;
    function GetActiveProject: Integer; safecall;
    procedure MoveLay(PageNo: Integer; Layer: Integer; NewLayer: Integer); safecall;
    procedure RemoveLay(PageNo: Integer; Layer: Integer); safecall;
    procedure SendOnTerminate(WinHandle: Integer; Msg: Integer); safecall;
    procedure CancelSaveCheck(Cancel: WordBool); safecall;
    procedure LockUserInterface(Happ: Integer; Lock: WordBool); safecall;
    procedure SetTextData(PageNo: Integer; TextHandle: Integer; AWidth: Integer; AHeight: Integer;
                          AOrigin: Integer; AOutLine: Integer; AColor: Integer; ABold: Integer;
                          AVisible: Integer; AAngle: Integer); safecall;
    procedure SetListOptions(ListType: Integer; PageNo: Integer; PrimSort: Integer;
                             SecSort: Integer; Options: Integer); safecall;
    procedure GetApplicationData(out Data: OleVariant); safecall;
    procedure SetApplicationData(Data: OleVariant); safecall;
    procedure GetProjectTitle(out ATitle: WideString); safecall;
    function GetMainWindowHandle: Integer; safecall;
    procedure NameGetSymbolHandles(APageType: Integer; AObjectType: Integer;
                                   const AName: WideString; out Handles: OleVariant); safecall;
    procedure SetSymbolConnPinNameData(PageNo: Integer; SymbolHandle: Integer;
                                       ConnectionIndex: Integer; AWidth: Integer; AHeight: Integer;
                                       AOrigin: Integer; AOutLine: Integer; AColor: Integer;
                                       ABold: Integer; AVisible: Integer; AAngle: Integer); safecall;
    procedure JoinParkedLines(PageNo: Integer); safecall;
    procedure SetSymbolScale(PageNo: Integer; SymbolHandle: Integer; AScale: Double); safecall;
    procedure RemovePage(PageNo: Integer); safecall;
    procedure GetAreas(PageNo: Integer; Layer: Integer; APenIndex: Integer; AMode: Integer;
                       out AList: OleVariant); safecall;
    function PlacePicture(const AFilename: WideString; Page: Integer; X: Integer; Y: Integer;
                          Z: Integer; Width: Integer; Height: Integer): Integer; safecall;
    procedure RemovePicture(PageNo: Integer; PictureHandle: Integer); safecall;
    function GetPictureCount(PageNo: Integer): Integer; safecall;
    function GetPictureHandle(PageNo: Integer; PictureIndex: Integer): Integer; safecall;
    procedure GetHeaderFileName(PageNo: Integer; out AFilename: WideString); safecall;
    procedure GetPageParms(PageNo: Integer; out ASize: Integer; out AType: Integer;
                           out AFunction: Integer; out AScale: Double); safecall;
    function GetSymbolDataFields(PageNo: Integer; SymbolHandle: Integer; out Fields: OleVariant;
                                 out Data: OleVariant): WordBool; safecall;
    function SetSymbolDataFields(PageNo: Integer; SymbolHandle: Integer; Data: OleVariant): WordBool; safecall;
    function CreateCableList(SortIndex: Integer; Lay: Integer; const StartPage: WideString;
                             const EndPage: WideString; Options: Integer): WordBool; safecall;
    procedure FreeCableList; safecall;
    function GetCableListCount: Integer; safecall;
    procedure GetCableList(ListIndex: Integer; out AList: OleVariant); safecall;
    procedure GetTermListData(ListIndex: Integer; DataType: Integer; out Symbol: Integer;
                              out Connection: Integer); safecall;
    procedure GetCableListData(ListIndex: Integer; DataType: Integer; out Symbol: Integer;
                               out Connection: Integer); safecall;
    procedure ActivateExtPopup(WinHandle: Integer; Msg: Integer); safecall;
    procedure ShowPopupMenu(ExtMenu: WordBool); safecall;
    procedure GetSelectionStatus(out AStatus: Integer); safecall;
    procedure GetSelection(AType: Integer; out AList: OleVariant); safecall;
    procedure SetPageTitle(PageNo: Integer; const ATitle: WideString); safecall;
    procedure ShowPickMenu(ShowIt: WordBool); safecall;
    procedure GetPopupMenuPos(InPixels: WordBool; out X: Integer; out Y: Integer); safecall;
    procedure SetPageExclIndexList(PageNo: Integer; Excl: WordBool); safecall;
    procedure SetElectricalLines(Electr: WordBool); safecall;
    function GetCurrentPageIndex: Integer; safecall;
    procedure SetPageType(PageNo: Integer; AType: Integer); safecall;
    procedure ClearSelection; safecall;
    procedure CreateNet(PageNo: Integer; SymbolHandle: Integer; ConnevtionIndex: Integer); safecall;
    procedure GetSelectionArea(out Left: Integer; out Right: Integer; out Bottom: Integer;
                               out Top: Integer; out AreaSelect: WordBool); safecall;
    procedure CreateCompList; safecall;
    procedure FreeCompList; safecall;
    function GetCompListCount: Integer; safecall;
    procedure CreatePartList; safecall;
    procedure FreePartList; safecall;
    function GetPartListCount: Integer; safecall;
    function GetDataField(const AName: WideString; ADataNo: Integer; ADataStatus: Integer;
                          AIndex: Integer): WideString; safecall;
    procedure ConnectSymbol(PageNo: Integer; SymbolHandle: Integer); safecall;
    function PlaceCableWire(PageNo: Integer; CableHandle: Integer; DownRight: WordBool; X: Integer;
                            Y: Integer; Z: Integer): Integer; safecall;
    procedure RemoveLine(PageNo: Integer; LineHandle: Integer); safecall;
    function GetLinePointCount(PageNo: Integer; LineHandle: Integer): Integer; safecall;
    function GetLinePointJoinIndex(PageNo: Integer; LineHandle: Integer; PointIndex: Integer): Integer; safecall;
    function GetSymbolConnectionJoinIndex(PageNo: Integer; SymbolHandle: Integer;
                                          ConnectionIndex: Integer): Integer; safecall;
    procedure ChangeLay(Layer: Integer); safecall;
    procedure ShowMsg(const AMsg: WideString; const BtnText: WideString; AType: Integer;
                      WinHandle: Integer; WinMsg: Integer); safecall;
    procedure CloseMsg; safecall;
    function InsertMenuItem(const AOwnerMenu: WideString; const ACaption: WideString;
                            AIndex: Integer; WinHandle: Integer; CmdID: Integer): Integer; safecall;
    function GetSymbolConnectionJoined(PageNo: Integer; SymbolHandle: Integer;
                                       ConnectionIndex: Integer): WordBool; safecall;
    procedure FindFirstNetInputOutput(AStatus: Integer; PageNo: Integer; SymbolHandle: Integer;
                                      ConnectionIndex: Integer; out PgNo: Integer;
                                      out SymHandle: Integer; out ConIndex: Integer); safecall;
    procedure FindFirstNetSymbolType(AType: Integer; PageNo: Integer; SymbolHandle: Integer;
                                     ConnectionIndex: Integer; out PgNo: Integer;
                                     out SymHandle: Integer; out ConIndex: Integer); safecall;
    function DataBaseMenu(const Articles: WideString; out AArticle: WideString;
                          out AType: WideString; out AFunction: WideString): WordBool; safecall;
    procedure SendOnSelection(WinHandle: Integer; Msg: Integer); safecall;
    function HighLightNet(PageNo: Integer; SymbolHandle: Integer; ConnectionIndex: Integer;
                          AColor: Integer): WordBool; safecall;
    function HighLightSymbol(PageNo: Integer; SymbolHandle: Integer; AColor: Integer): WordBool; safecall;
    function FileNotSaved(PageNo: Integer): WordBool; safecall;
    function SetSymbolQuantity(PageNo: Integer; SymbolHandle: Integer; Quantity: Double): WordBool; safecall;
    function SetSymbolConnTextData(PageNo: Integer; SymbolHandle: Integer;
                                   ConnectionIndex: Integer; TextIndex: Integer; AWidth: Integer;
                                   AHeight: Integer; AOrigin: Integer; AOutLine: Integer;
                                   AColor: Integer; ABold: Integer; AVisible: Integer;
                                   AAngle: Integer): WordBool; safecall;
    function GetNetListExt(NetIndex: Integer; Flags: Integer; out AList: OleVariant): WordBool; safecall;
    function CopySymbol(SrcProjectNo: Integer; DstProjectNo: Integer; SrcPageNo: Integer;
                        DstPageNo: Integer; SrcSymbolHandle: Integer; X: Integer; Y: Integer;
                        Z: Integer; Layer: Integer): Integer; safecall;
    function PinSwop(PageNo: Integer; SymbolHandle: Integer; Connection1Index: Integer;
                     Connection2Index: Integer): WordBool; safecall;
    function CopyPage(SrcProjectNo: Integer; DstProjectNo: Integer; SrcPageNo: Integer;
                      DstPageNo: Integer): WordBool; safecall;
    function CopyEmptyPage(SrcProjectNo: Integer; DstProjectNo: Integer; SrcPageNo: Integer;
                           DstPageNo: Integer): WordBool; safecall;
    function SetOleProject(ProjectNo: Integer): WordBool; safecall;
    function InsertProject(PageNo: Integer; const AFilename: WideString;
                           const APageNumber: WideString; Options: Integer): WordBool; safecall;
    function SetPageNumber(PageNo: Integer; const ANumber: WideString): WordBool; safecall;
    function GetPageNumber(PageNo: Integer; out ANumber: WideString): WordBool; safecall;
    function SetPageTagData(PageNo: Integer; AGroupNo: Integer; AList: OleVariant): WordBool; safecall;
    function GetPageTagData(PageNo: Integer; AGroupNo: Integer; out AList: OleVariant): WordBool; safecall;
    function GetPageTagNames(PageNo: Integer; AGroupNo: Integer; out AList: OleVariant): WordBool; safecall;
    function CreatePageTags(PageNo: Integer): WordBool; safecall;
    function GetInsertionProints(PageNo: Integer; MinDist: Integer; out X1: Integer;
                                 out Y1: Integer; out Z1: Integer; out X2: Integer;
                                 out Y2: Integer; out Z2: Integer): Integer; safecall;
    function LoadSubDrawing(const AFilename: WideString): WordBool; safecall;
    function ReleaseSubDrawing: WordBool; safecall;
    function PlaceSubDrawing(PageNo: Integer; X: Integer; Y: Integer; Z: Integer; Layer: Integer;
                             Group: Integer; Options: Integer): WordBool; safecall;
    function GetPageData(PageNo: Integer; out Defs: OleVariant; out Data: OleVariant): WordBool; safecall;
    function SetPageData(PageNo: Integer; Data: OleVariant): WordBool; safecall;
    function MirrorSymbol(PageNo: Integer; SymbolHandle: Integer; Vertical: WordBool;
                          FromCurDir: WordBool): WordBool; safecall;
    function OpenProjectTemplate(const AFilename: WideString): Integer; safecall;
    function GetTextCount(PageNo: Integer): Integer; safecall;
    function GetTextHandle(PageNo: Integer; TextIndex: Integer): Integer; safecall;
    function GetText(PageNo: Integer; TextHandle: Integer; out AText: WideString): WordBool; safecall;
    function SetText(PageNo: Integer; TextHandle: Integer; const AText: WideString): WordBool; safecall;
    function CreateConnList(ListType: Integer; Layer: Integer): WordBool; safecall;
    function FreeConnList: WordBool; safecall;
    function GetConnListCount: Integer; safecall;
    function GetConnList(ConnIndex: Integer; Flags: Integer; out AList: OleVariant): WordBool; safecall;
    procedure SendOnProject(WndHandle: Integer; Msg: Integer); safecall;
    function SetCursor(ACursor: Integer): Integer; safecall;
    function ApplicationDataReadSection(PageNo: Integer; const Section: WideString;
                                        out Sata: OleVariant): WordBool; safecall;
    function ApplicationDataReadString(PageNo: Integer; const Section: WideString;
                                       const Ident: WideString; const Default: WideString;
                                       out Value: WideString): WordBool; safecall;
    function ApplicationDataWriteSection(PageNo: Integer; const Section: WideString;
                                         Data: OleVariant): WordBool; safecall;
    function ApplicationDataWriteString(PageNo: Integer; const Section: WideString;
                                        const Ident: WideString; const Value: WideString): WordBool; safecall;
    function GetReferenceID(PageNo: Integer; SymbolHandle: Integer; out AFunction: WideString;
                            out ALocation: WideString; out AName: WideString;
                            out ASubName: WideString): WordBool; safecall;
    function SetReferenceID(PageNo: Integer; SymbolHandle: Integer; const AFunction: WideString;
                            const ALocation: WideString; const AName: WideString;
                            const ASubName: WideString): WordBool; safecall;
    function CreateReferenceID(AType: Integer; const AID: WideString; const ADiscr: WideString): WordBool; safecall;
    function UpdateProjectFromDB(PageNo: Integer; SymbolHandle: Integer): WordBool; safecall;
    function GetSymbolCompGroup(PageNo: Integer; SymbolHandle: Integer; out AGroupNo: Integer;
                                out APosNo: Integer): WordBool; safecall;
    function CreateSymbolDataField(const AName: WideString): WordBool; safecall;
    function GetPageUsedDataFields(AType: Integer; PageNo: Integer; out Fields: OleVariant): WordBool; safecall;
    function PlaceCableScreen(PageNo: Integer; CableHandle: Integer): Integer; safecall;
    function GetPageTemplateNames(PageNo: Integer; out AList: OleVariant): WordBool; safecall;
    function GetSymbolStates(PageNo: Integer; SymbolHandle: Integer; out States: OleVariant): WordBool; safecall;
    function SetSymbolState(PageNo: Integer; SymbolHandle: Integer; AState: Integer): WordBool; safecall;
    function GetSymbolState(PageNo: Integer; SymbolHandle: Integer; out AState: Integer): WordBool; safecall;
    function Print(PageNo: Integer): WordBool; safecall;
    function LoadMechPageCompList(PageNo: Integer): WordBool; safecall;
    procedure SendOnDesignSymbol(WndHandle: Integer; Msg: Integer); safecall;
    function SortCreatedList(AListType: Integer; APrimSort: Integer; ASecSort: Integer): WordBool; safecall;
    function SetObjectVisible(APageNo: Integer; ObjectType: Integer; ObjectHandle: Integer;
                              Visible: WordBool): WordBool; safecall;
    function SetProjectTitle(const ATitle: WideString): WordBool; safecall;
    function SetSymbolConnectionPolygon(PageNo: Integer; SymbolHandle: Integer;
                                        ConnectionIndex: Integer; AShape: Integer;
                                        BordorColor: Integer; FillColor: Integer;
                                        const Coors: OleVariant; Options: Integer): WordBool; safecall;
    function EXCreateCompList(const ParamStr: WideString): WordBool; safecall;
    function EXCreatePartList(const ParamStr: WideString): WordBool; safecall;
    function SymbolApplDataReadString(PageNo: Integer; SymbolHandle: Integer;
                                      const Section: WideString; const Ident: WideString;
                                      const Default: WideString; out Value: WideString): WordBool; safecall;
    function SymbolApplDataWriteString(PageNo: Integer; SymbolHandle: Integer;
                                       const Section: WideString; const Ident: WideString;
                                       const Value: WideString): WordBool; safecall;
    function GetDBFieldValue(const SearchField: WideString; const SearchValue: WideString;
                             const FieldName: WideString; out Value: WideString): WordBool; safecall;
  end;

// *********************************************************************//
// DispIntf:  IAppDisp
// Flags:     (4432) Hidden Dual OleAutomation Dispatchable
// GUID:      {C5E8FDA5-C88D-11D1-AE4B-0020AF16D64A}
// *********************************************************************//
  IAppDisp = dispinterface
    ['{C5E8FDA5-C88D-11D1-AE4B-0020AF16D64A}']
    function CreateProject(const AProjectTitle: WideString; const AFilename: WideString): Integer; dispid 1;
    function OpenProject(const AFilename: WideString): Integer; dispid 59;
    procedure CloseProject(ProjectNo: Integer; SavedCheck: WordBool); dispid 49;
    procedure SaveWorkFile(ProjectNo: Integer; const AFilename: WideString); dispid 40;
    procedure GetProjectData(out Defs: OleVariant; out Data: OleVariant); dispid 41;
    procedure SetProjectData(Data: OleVariant); dispid 42;
    function OpenSymbol(const AFilename: WideString): Integer; dispid 73;
    procedure UpdateList(ListType: Integer); dispid 48;
    procedure CreateNetList(ListType: Integer; Layer: Integer); dispid 50;
    procedure FreeNetList; dispid 51;
    function GetNetListCount: Integer; dispid 52;
    procedure GetNetList(NetIndex: Integer; out AList: OleVariant); dispid 53;
    procedure CreateDynCableList(Layer: Integer); dispid 55;
    procedure FreeDynCableList; dispid 56;
    function GetDynCableListCount: Integer; dispid 57;
    procedure GetDynCableList(ListIndex: Integer; out AList: OleVariant); dispid 58;
    procedure NewPage(PageNo: Integer); dispid 15;
    procedure SetPageFunction(PageNo: Integer; AFunction: Integer); dispid 54;
    procedure GetPageBox(PageNo: Integer; out X1: Integer; out Y1: Integer; out X2: Integer;
                         out Y2: Integer); dispid 2;
    function GetPageCount: Integer; dispid 60;
    function PlaceSymbol(const SymbolFilename: WideString; Page: Integer; X: Integer; Y: Integer;
                         Z: Integer; Layer: Integer): Integer; dispid 3;
    function GetSymbolCount(PageNo: Integer): Integer; dispid 46;
    function GetSymbolHandle(PageNo: Integer; SymbolIndex: Integer): Integer; dispid 47;
    procedure RemoveSymbol(PageNo: Integer; SymbolHandle: Integer); dispid 4;
    procedure SetSymbolText(PageNo: Integer; SymbolHandle: Integer; const Name: WideString;
                            const TypeText: WideString; const Article: WideString;
                            const FunctionText: WideString); dispid 5;
    procedure SetSymbolType(PageNo: Integer; SymbolHandle: Integer; AType: Integer); dispid 39;
    function GetSymbolType(PageNo: Integer; SymbolHandle: Integer; out AType: Integer): WordBool; dispid 61;
    procedure GetSymbolText(PageNo: Integer; SymbolHandle: Integer; out Name: WideString;
                            out TypeText: WideString; out Article: WideString;
                            out FunctionText: WideString); dispid 6;
    procedure GetSymbolBox(PageNo: Integer; SymbolHandle: Integer; BoxType: Integer;
                           out X1: Integer; out Y1: Integer; out X2: Integer; out Y2: Integer); dispid 7;
    procedure MoveSymbol(PageNo: Integer; SymbolHandle: Integer; NewPage: Integer; X: Integer;
                         Y: Integer; Z: Integer; Layer: Integer); dispid 8;
    procedure GetSymbolPos(PageNo: Integer; SymbolHandle: Integer; out X: Integer; out Y: Integer;
                           out Z: Integer; out Layer: Integer); dispid 9;
    procedure GetSymbolConnectionPosRel(PageNo: Integer; SymbolHandle: Integer;
                                        ConnectionIndex: Integer; out X: Integer; out Y: Integer;
                                        out Z: Integer); dispid 10;
    procedure GetSymbolConnectionIOstatus(PageNo: Integer; SymbolHandle: Integer;
                                          ConnectionIndex: Integer; out AStatus: Integer); dispid 43;
    function PlaceLine(PageNo: Integer; Layer: Integer; StartX: Integer; StartY: Integer;
                       StartZ: Integer; EndX: Integer; EndY: Integer; EndZ: Integer): Integer; dispid 11;
    function PlacePolyLine(PageNo: Integer; Layer: Integer; AList: OleVariant): Integer; dispid 34;
    function GetLineCount(PageNo: Integer): Integer; dispid 44;
    function GetLineHandle(PageNo: Integer; LineIndex: Integer): Integer; dispid 45;
    procedure SetLineOptions(Color: Integer; Style: Integer; Width: Integer); dispid 12;
    procedure UpdateDrawing; dispid 13;
    function GetConnections(PageIndex: Integer; SymbolHandle: Integer): OleVariant; dispid 14;
    function PlaceDynamicCable(PageNo: Integer; Layer: Integer; Direction: Integer;
                               AList: OleVariant): Integer; dispid 16;
    function GetSymbolConnectionCount(PageNo: Integer; SymbolHandle: Integer): Integer; dispid 17;
    procedure CreateJoins(PageNo: Integer); dispid 18;
    function GetSymbolDataFieldCount(PageNo: Integer; SymbolHandle: Integer): Integer; dispid 19;
    procedure GetSymbolDataField(PageNo: Integer; SymbolHandle: Integer; FieldIndex: Integer;
                                 out AName: WideString; out AText: WideString); dispid 20;
    procedure SetSymbolDataField(PageNo: Integer; SymbolHandle: Integer; FieldIndex: Integer;
                                 const AText: WideString); dispid 21;
    procedure SetSymbolConnectionText(PageNo: Integer; SymbolHandle: Integer;
                                      ConnectionIndex: Integer; const Name: WideString;
                                      const FunctionText: WideString; const LabelText: WideString;
                                      const Description: WideString); dispid 22;
    procedure GetSymbolConnectionText(PageNo: Integer; SymbolHandle: Integer;
                                      ConnectionIndex: Integer; out Name: WideString;
                                      out FunctionText: WideString; out LabelText: WideString;
                                      out Description: WideString); dispid 23;
    function PlaceHeader(PageIndex: Integer; const HeaderFilename: WideString): WordBool; dispid 24;
    function PlaceDynamicSymbol(const SymbolFilename: WideString; PageNo: Integer; X: Integer;
                                Y: Integer; Z: Integer; Layer: Integer; AList: OleVariant): Integer; dispid 25;
    procedure GetSymbolFileName(PageNo: Integer; SymbolHandle: Integer; out AFilename: WideString); dispid 71;
    procedure GetSymbolTitle(PageNo: Integer; SymbolHandle: Integer; out ATitle: WideString); dispid 26;
    procedure SetSymbolNameData(PageNo: Integer; SymbolHandle: Integer; AWidth: Integer;
                                AHeight: Integer; AOrigin: Integer; AOutLine: Integer;
                                AColor: Integer; ABold: Integer; AVisible: Integer; AAngle: Integer); dispid 27;
    procedure SetSymbolNamePos(PageNo: Integer; SymbolHandle: Integer; X: Integer; Y: Integer;
                               Z: Integer); dispid 28;
    procedure SetSymbolTypeTextData(PageNo: Integer; SymbolHandle: Integer; AWidth: Integer;
                                    AHeight: Integer; AOrigin: Integer; AOutLine: Integer;
                                    AColor: Integer; ABold: Integer; AVisible: Integer;
                                    AAngle: Integer); dispid 35;
    procedure SetSymbolTypeTextPos(PageNo: Integer; SymbolHandle: Integer; X: Integer; Y: Integer;
                                   Z: Integer); dispid 36;
    procedure SetSymbolArticleData(PageNo: Integer; SymbolHandle: Integer; AWidth: Integer;
                                   AHeight: Integer; AOrigin: Integer; AOutLine: Integer;
                                   AColor: Integer; ABold: Integer; AVisible: Integer;
                                   AAngle: Integer); dispid 37;
    procedure SetSymbolArticlePos(PageNo: Integer; SymbolHandle: Integer; X: Integer; Y: Integer;
                                  Z: Integer); dispid 38;
    procedure SetSymbolFunctionData(PageNo: Integer; SymbolHandle: Integer; AWidth: Integer;
                                    AHeight: Integer; AOrigin: Integer; AOutLine: Integer;
                                    AColor: Integer; ABold: Integer; AVisible: Integer;
                                    AAngle: Integer); dispid 29;
    procedure SetSymbolFunctionPos(PageNo: Integer; SymbolHandle: Integer; X: Integer; Y: Integer;
                                   Z: Integer); dispid 30;
    procedure GetPageHeaders(out Headers: OleVariant); dispid 31;
    procedure GetSymbolFileNames(out FileNames: OleVariant); dispid 32;
    procedure GetTextWidth(ACharWidth: Integer; const AText: WideString; out ATextWidth: Integer); dispid 33;
    function AddSpeedButton(const ABarName: WideString; const AHint: WideString;
                            const Caption_BMP: WideString; WinHandle: Integer; CmdID: Integer): Integer; dispid 66;
    function AddMenuItem(const AOwnerMenu: WideString; const ACaption: WideString;
                         WinHandle: Integer; CmdID: Integer): Integer; dispid 67;
    function RemoveCmdItem(AItem: Integer): WordBool; dispid 68;
    function EnableCmdItem(AItem: Integer; AVisible: WordBool; AEnabled: WordBool): WordBool; dispid 69;
    procedure CreateTermList; dispid 62;
    procedure FreeTermList; dispid 63;
    function GetTermListCount: Integer; dispid 64;
    procedure GetTermList(ListIndex: Integer; out AList: OleVariant); dispid 65;
    procedure BringToFront; dispid 70;
    function GetProgramState: Integer; dispid 72;
    function AddMenu(const AOwnerMenu: WideString; const ACaption: WideString;
                     out AMenuName: WideString): Integer; dispid 74;
    procedure SetPageReference(PageNo: Integer; AFlags: Integer); dispid 75;
    procedure SetSymbolConnTextsVisible(PageNo: Integer; SymbolHandle: Integer; ATexts: Integer;
                                        AVisible: Integer); dispid 76;
    function PlaceText(const AText: WideString; Page: Integer; X: Integer; Y: Integer; Z: Integer;
                       Layer: Integer): Integer; dispid 77;
    procedure SetLayVisible(Layer: Integer; AVisible: Integer); dispid 78;
    procedure SymbolMenu(const ALibrary: WideString); dispid 79;
    procedure DontShutDownOnRelease; dispid 80;
    function GetProgramVersion: WideString; dispid 81;
    function GetProgramEnv(const EnvField: WideString): WideString; dispid 82;
    procedure UpdateMenus; dispid 83;
    procedure RotateSymbol(PageNo: Integer; SymbolHandle: Integer; Angle: Integer); dispid 84;
    procedure SetProjectTextsVisible(ATexts: Integer; AVisible: Integer); dispid 85;
    procedure TransferLibType(SrcProjectNo: Integer; DstProjectNo: Integer;
                              const AFilename: WideString); dispid 86;
    function GetActiveProjectHandle: Integer; dispid 87;
    function GetFunctionID(const FunctionName: WideString): Integer; dispid 88;
    procedure FunctionActivate(AID: Integer; Activate: WordBool); dispid 89;
    procedure FunctionVisible(AID: Integer; Visible: WordBool); dispid 90;
    procedure FunctionEnable(AID: Integer; Enable: WordBool); dispid 91;
    procedure SetFunctionMsg(AID: Integer; WinHandle: Integer; Msg: Integer); dispid 92;
    procedure GetProjectName(out AFilename: WideString); dispid 93;
    function GetActiveProject: Integer; dispid 94;
    procedure MoveLay(PageNo: Integer; Layer: Integer; NewLayer: Integer); dispid 95;
    procedure RemoveLay(PageNo: Integer; Layer: Integer); dispid 96;
    procedure SendOnTerminate(WinHandle: Integer; Msg: Integer); dispid 97;
    procedure CancelSaveCheck(Cancel: WordBool); dispid 98;
    procedure LockUserInterface(Happ: Integer; Lock: WordBool); dispid 99;
    procedure SetTextData(PageNo: Integer; TextHandle: Integer; AWidth: Integer; AHeight: Integer;
                          AOrigin: Integer; AOutLine: Integer; AColor: Integer; ABold: Integer;
                          AVisible: Integer; AAngle: Integer); dispid 100;
    procedure SetListOptions(ListType: Integer; PageNo: Integer; PrimSort: Integer;
                             SecSort: Integer; Options: Integer); dispid 101;
    procedure GetApplicationData(out Data: OleVariant); dispid 102;
    procedure SetApplicationData(Data: OleVariant); dispid 103;
    procedure GetProjectTitle(out ATitle: WideString); dispid 104;
    function GetMainWindowHandle: Integer; dispid 105;
    procedure NameGetSymbolHandles(APageType: Integer; AObjectType: Integer;
                                   const AName: WideString; out Handles: OleVariant); dispid 106;
    procedure SetSymbolConnPinNameData(PageNo: Integer; SymbolHandle: Integer;
                                       ConnectionIndex: Integer; AWidth: Integer; AHeight: Integer;
                                       AOrigin: Integer; AOutLine: Integer; AColor: Integer;
                                       ABold: Integer; AVisible: Integer; AAngle: Integer); dispid 107;
    procedure JoinParkedLines(PageNo: Integer); dispid 108;
    procedure SetSymbolScale(PageNo: Integer; SymbolHandle: Integer; AScale: Double); dispid 109;
    procedure RemovePage(PageNo: Integer); dispid 110;
    procedure GetAreas(PageNo: Integer; Layer: Integer; APenIndex: Integer; AMode: Integer;
                       out AList: OleVariant); dispid 111;
    function PlacePicture(const AFilename: WideString; Page: Integer; X: Integer; Y: Integer;
                          Z: Integer; Width: Integer; Height: Integer): Integer; dispid 112;
    procedure RemovePicture(PageNo: Integer; PictureHandle: Integer); dispid 113;
    function GetPictureCount(PageNo: Integer): Integer; dispid 114;
    function GetPictureHandle(PageNo: Integer; PictureIndex: Integer): Integer; dispid 115;
    procedure GetHeaderFileName(PageNo: Integer; out AFilename: WideString); dispid 116;
    procedure GetPageParms(PageNo: Integer; out ASize: Integer; out AType: Integer;
                           out AFunction: Integer; out AScale: Double); dispid 117;
    function GetSymbolDataFields(PageNo: Integer; SymbolHandle: Integer; out Fields: OleVariant;
                                 out Data: OleVariant): WordBool; dispid 118;
    function SetSymbolDataFields(PageNo: Integer; SymbolHandle: Integer; Data: OleVariant): WordBool; dispid 119;
    function CreateCableList(SortIndex: Integer; Lay: Integer; const StartPage: WideString;
                             const EndPage: WideString; Options: Integer): WordBool; dispid 120;
    procedure FreeCableList; dispid 121;
    function GetCableListCount: Integer; dispid 122;
    procedure GetCableList(ListIndex: Integer; out AList: OleVariant); dispid 123;
    procedure GetTermListData(ListIndex: Integer; DataType: Integer; out Symbol: Integer;
                              out Connection: Integer); dispid 124;
    procedure GetCableListData(ListIndex: Integer; DataType: Integer; out Symbol: Integer;
                               out Connection: Integer); dispid 125;
    procedure ActivateExtPopup(WinHandle: Integer; Msg: Integer); dispid 126;
    procedure ShowPopupMenu(ExtMenu: WordBool); dispid 127;
    procedure GetSelectionStatus(out AStatus: Integer); dispid 128;
    procedure GetSelection(AType: Integer; out AList: OleVariant); dispid 129;
    procedure SetPageTitle(PageNo: Integer; const ATitle: WideString); dispid 130;
    procedure ShowPickMenu(ShowIt: WordBool); dispid 131;
    procedure GetPopupMenuPos(InPixels: WordBool; out X: Integer; out Y: Integer); dispid 132;
    procedure SetPageExclIndexList(PageNo: Integer; Excl: WordBool); dispid 133;
    procedure SetElectricalLines(Electr: WordBool); dispid 134;
    function GetCurrentPageIndex: Integer; dispid 135;
    procedure SetPageType(PageNo: Integer; AType: Integer); dispid 136;
    procedure ClearSelection; dispid 137;
    procedure CreateNet(PageNo: Integer; SymbolHandle: Integer; ConnevtionIndex: Integer); dispid 138;
    procedure GetSelectionArea(out Left: Integer; out Right: Integer; out Bottom: Integer;
                               out Top: Integer; out AreaSelect: WordBool); dispid 139;
    procedure CreateCompList; dispid 140;
    procedure FreeCompList; dispid 141;
    function GetCompListCount: Integer; dispid 142;
    procedure CreatePartList; dispid 143;
    procedure FreePartList; dispid 144;
    function GetPartListCount: Integer; dispid 145;
    function GetDataField(const AName: WideString; ADataNo: Integer; ADataStatus: Integer;
                          AIndex: Integer): WideString; dispid 146;
    procedure ConnectSymbol(PageNo: Integer; SymbolHandle: Integer); dispid 147;
    function PlaceCableWire(PageNo: Integer; CableHandle: Integer; DownRight: WordBool; X: Integer;
                            Y: Integer; Z: Integer): Integer; dispid 148;
    procedure RemoveLine(PageNo: Integer; LineHandle: Integer); dispid 149;
    function GetLinePointCount(PageNo: Integer; LineHandle: Integer): Integer; dispid 150;
    function GetLinePointJoinIndex(PageNo: Integer; LineHandle: Integer; PointIndex: Integer): Integer; dispid 151;
    function GetSymbolConnectionJoinIndex(PageNo: Integer; SymbolHandle: Integer;
                                          ConnectionIndex: Integer): Integer; dispid 152;
    procedure ChangeLay(Layer: Integer); dispid 153;
    procedure ShowMsg(const AMsg: WideString; const BtnText: WideString; AType: Integer;
                      WinHandle: Integer; WinMsg: Integer); dispid 154;
    procedure CloseMsg; dispid 155;
    function InsertMenuItem(const AOwnerMenu: WideString; const ACaption: WideString;
                            AIndex: Integer; WinHandle: Integer; CmdID: Integer): Integer; dispid 156;
    function GetSymbolConnectionJoined(PageNo: Integer; SymbolHandle: Integer;
                                       ConnectionIndex: Integer): WordBool; dispid 157;
    procedure FindFirstNetInputOutput(AStatus: Integer; PageNo: Integer; SymbolHandle: Integer;
                                      ConnectionIndex: Integer; out PgNo: Integer;
                                      out SymHandle: Integer; out ConIndex: Integer); dispid 158;
    procedure FindFirstNetSymbolType(AType: Integer; PageNo: Integer; SymbolHandle: Integer;
                                     ConnectionIndex: Integer; out PgNo: Integer;
                                     out SymHandle: Integer; out ConIndex: Integer); dispid 159;
    function DataBaseMenu(const Articles: WideString; out AArticle: WideString;
                          out AType: WideString; out AFunction: WideString): WordBool; dispid 160;
    procedure SendOnSelection(WinHandle: Integer; Msg: Integer); dispid 161;
    function HighLightNet(PageNo: Integer; SymbolHandle: Integer; ConnectionIndex: Integer;
                          AColor: Integer): WordBool; dispid 162;
    function HighLightSymbol(PageNo: Integer; SymbolHandle: Integer; AColor: Integer): WordBool; dispid 163;
    function FileNotSaved(PageNo: Integer): WordBool; dispid 164;
    function SetSymbolQuantity(PageNo: Integer; SymbolHandle: Integer; Quantity: Double): WordBool; dispid 165;
    function SetSymbolConnTextData(PageNo: Integer; SymbolHandle: Integer;
                                   ConnectionIndex: Integer; TextIndex: Integer; AWidth: Integer;
                                   AHeight: Integer; AOrigin: Integer; AOutLine: Integer;
                                   AColor: Integer; ABold: Integer; AVisible: Integer;
                                   AAngle: Integer): WordBool; dispid 166;
    function GetNetListExt(NetIndex: Integer; Flags: Integer; out AList: OleVariant): WordBool; dispid 167;
    function CopySymbol(SrcProjectNo: Integer; DstProjectNo: Integer; SrcPageNo: Integer;
                        DstPageNo: Integer; SrcSymbolHandle: Integer; X: Integer; Y: Integer;
                        Z: Integer; Layer: Integer): Integer; dispid 168;
    function PinSwop(PageNo: Integer; SymbolHandle: Integer; Connection1Index: Integer;
                     Connection2Index: Integer): WordBool; dispid 169;
    function CopyPage(SrcProjectNo: Integer; DstProjectNo: Integer; SrcPageNo: Integer;
                      DstPageNo: Integer): WordBool; dispid 170;
    function CopyEmptyPage(SrcProjectNo: Integer; DstProjectNo: Integer; SrcPageNo: Integer;
                           DstPageNo: Integer): WordBool; dispid 171;
    function SetOleProject(ProjectNo: Integer): WordBool; dispid 172;
    function InsertProject(PageNo: Integer; const AFilename: WideString;
                           const APageNumber: WideString; Options: Integer): WordBool; dispid 173;
    function SetPageNumber(PageNo: Integer; const ANumber: WideString): WordBool; dispid 174;
    function GetPageNumber(PageNo: Integer; out ANumber: WideString): WordBool; dispid 175;
    function SetPageTagData(PageNo: Integer; AGroupNo: Integer; AList: OleVariant): WordBool; dispid 176;
    function GetPageTagData(PageNo: Integer; AGroupNo: Integer; out AList: OleVariant): WordBool; dispid 177;
    function GetPageTagNames(PageNo: Integer; AGroupNo: Integer; out AList: OleVariant): WordBool; dispid 178;
    function CreatePageTags(PageNo: Integer): WordBool; dispid 179;
    function GetInsertionProints(PageNo: Integer; MinDist: Integer; out X1: Integer;
                                 out Y1: Integer; out Z1: Integer; out X2: Integer;
                                 out Y2: Integer; out Z2: Integer): Integer; dispid 181;
    function LoadSubDrawing(const AFilename: WideString): WordBool; dispid 182;
    function ReleaseSubDrawing: WordBool; dispid 183;
    function PlaceSubDrawing(PageNo: Integer; X: Integer; Y: Integer; Z: Integer; Layer: Integer;
                             Group: Integer; Options: Integer): WordBool; dispid 184;
    function GetPageData(PageNo: Integer; out Defs: OleVariant; out Data: OleVariant): WordBool; dispid 180;
    function SetPageData(PageNo: Integer; Data: OleVariant): WordBool; dispid 185;
    function MirrorSymbol(PageNo: Integer; SymbolHandle: Integer; Vertical: WordBool;
                          FromCurDir: WordBool): WordBool; dispid 186;
    function OpenProjectTemplate(const AFilename: WideString): Integer; dispid 187;
    function GetTextCount(PageNo: Integer): Integer; dispid 188;
    function GetTextHandle(PageNo: Integer; TextIndex: Integer): Integer; dispid 189;
    function GetText(PageNo: Integer; TextHandle: Integer; out AText: WideString): WordBool; dispid 190;
    function SetText(PageNo: Integer; TextHandle: Integer; const AText: WideString): WordBool; dispid 191;
    function CreateConnList(ListType: Integer; Layer: Integer): WordBool; dispid 192;
    function FreeConnList: WordBool; dispid 193;
    function GetConnListCount: Integer; dispid 194;
    function GetConnList(ConnIndex: Integer; Flags: Integer; out AList: OleVariant): WordBool; dispid 195;
    procedure SendOnProject(WndHandle: Integer; Msg: Integer); dispid 196;
    function SetCursor(ACursor: Integer): Integer; dispid 197;
    function ApplicationDataReadSection(PageNo: Integer; const Section: WideString;
                                        out Sata: OleVariant): WordBool; dispid 198;
    function ApplicationDataReadString(PageNo: Integer; const Section: WideString;
                                       const Ident: WideString; const Default: WideString;
                                       out Value: WideString): WordBool; dispid 199;
    function ApplicationDataWriteSection(PageNo: Integer; const Section: WideString;
                                         Data: OleVariant): WordBool; dispid 200;
    function ApplicationDataWriteString(PageNo: Integer; const Section: WideString;
                                        const Ident: WideString; const Value: WideString): WordBool; dispid 201;
    function GetReferenceID(PageNo: Integer; SymbolHandle: Integer; out AFunction: WideString;
                            out ALocation: WideString; out AName: WideString;
                            out ASubName: WideString): WordBool; dispid 202;
    function SetReferenceID(PageNo: Integer; SymbolHandle: Integer; const AFunction: WideString;
                            const ALocation: WideString; const AName: WideString;
                            const ASubName: WideString): WordBool; dispid 203;
    function CreateReferenceID(AType: Integer; const AID: WideString; const ADiscr: WideString): WordBool; dispid 204;
    function UpdateProjectFromDB(PageNo: Integer; SymbolHandle: Integer): WordBool; dispid 205;
    function GetSymbolCompGroup(PageNo: Integer; SymbolHandle: Integer; out AGroupNo: Integer;
                                out APosNo: Integer): WordBool; dispid 206;
    function CreateSymbolDataField(const AName: WideString): WordBool; dispid 207;
    function GetPageUsedDataFields(AType: Integer; PageNo: Integer; out Fields: OleVariant): WordBool; dispid 208;
    function PlaceCableScreen(PageNo: Integer; CableHandle: Integer): Integer; dispid 209;
    function GetPageTemplateNames(PageNo: Integer; out AList: OleVariant): WordBool; dispid 210;
    function GetSymbolStates(PageNo: Integer; SymbolHandle: Integer; out States: OleVariant): WordBool; dispid 211;
    function SetSymbolState(PageNo: Integer; SymbolHandle: Integer; AState: Integer): WordBool; dispid 212;
    function GetSymbolState(PageNo: Integer; SymbolHandle: Integer; out AState: Integer): WordBool; dispid 213;
    function Print(PageNo: Integer): WordBool; dispid 214;
    function LoadMechPageCompList(PageNo: Integer): WordBool; dispid 215;
    procedure SendOnDesignSymbol(WndHandle: Integer; Msg: Integer); dispid 216;
    function SortCreatedList(AListType: Integer; APrimSort: Integer; ASecSort: Integer): WordBool; dispid 217;
    function SetObjectVisible(APageNo: Integer; ObjectType: Integer; ObjectHandle: Integer;
                              Visible: WordBool): WordBool; dispid 218;
    function SetProjectTitle(const ATitle: WideString): WordBool; dispid 219;
    function SetSymbolConnectionPolygon(PageNo: Integer; SymbolHandle: Integer;
                                        ConnectionIndex: Integer; AShape: Integer;
                                        BordorColor: Integer; FillColor: Integer;
                                        const Coors: OleVariant; Options: Integer): WordBool; dispid 220;
    function EXCreateCompList(const ParamStr: WideString): WordBool; dispid 221;
    function EXCreatePartList(const ParamStr: WideString): WordBool; dispid 222;
    function SymbolApplDataReadString(PageNo: Integer; SymbolHandle: Integer;
                                      const Section: WideString; const Ident: WideString;
                                      const Default: WideString; out Value: WideString): WordBool; dispid 223;
    function SymbolApplDataWriteString(PageNo: Integer; SymbolHandle: Integer;
                                       const Section: WideString; const Ident: WideString;
                                       const Value: WideString): WordBool; dispid 224;
    function GetDBFieldValue(const SearchField: WideString; const SearchValue: WideString;
                             const FieldName: WideString; out Value: WideString): WordBool; dispid 225;
  end;

// *********************************************************************//
// Interface: IPCsApplicationObsolete
// Flags:     (4432) Hidden Dual OleAutomation Dispatchable
// GUID:      {6CF6C0C6-D8FB-11D5-9030-0020AF16D6BF}
// *********************************************************************//
  IPCsApplicationObsolete = interface(IApp)
    ['{6CF6C0C6-D8FB-11D5-9030-0020AF16D6BF}']
  end;

// *********************************************************************//
// DispIntf:  IPCsApplicationObsoleteDisp
// Flags:     (4432) Hidden Dual OleAutomation Dispatchable
// GUID:      {6CF6C0C6-D8FB-11D5-9030-0020AF16D6BF}
// *********************************************************************//
  IPCsApplicationObsoleteDisp = dispinterface
    ['{6CF6C0C6-D8FB-11D5-9030-0020AF16D6BF}']
    function CreateProject(const AProjectTitle: WideString; const AFilename: WideString): Integer; dispid 1;
    function OpenProject(const AFilename: WideString): Integer; dispid 59;
    procedure CloseProject(ProjectNo: Integer; SavedCheck: WordBool); dispid 49;
    procedure SaveWorkFile(ProjectNo: Integer; const AFilename: WideString); dispid 40;
    procedure GetProjectData(out Defs: OleVariant; out Data: OleVariant); dispid 41;
    procedure SetProjectData(Data: OleVariant); dispid 42;
    function OpenSymbol(const AFilename: WideString): Integer; dispid 73;
    procedure UpdateList(ListType: Integer); dispid 48;
    procedure CreateNetList(ListType: Integer; Layer: Integer); dispid 50;
    procedure FreeNetList; dispid 51;
    function GetNetListCount: Integer; dispid 52;
    procedure GetNetList(NetIndex: Integer; out AList: OleVariant); dispid 53;
    procedure CreateDynCableList(Layer: Integer); dispid 55;
    procedure FreeDynCableList; dispid 56;
    function GetDynCableListCount: Integer; dispid 57;
    procedure GetDynCableList(ListIndex: Integer; out AList: OleVariant); dispid 58;
    procedure NewPage(PageNo: Integer); dispid 15;
    procedure SetPageFunction(PageNo: Integer; AFunction: Integer); dispid 54;
    procedure GetPageBox(PageNo: Integer; out X1: Integer; out Y1: Integer; out X2: Integer;
                         out Y2: Integer); dispid 2;
    function GetPageCount: Integer; dispid 60;
    function PlaceSymbol(const SymbolFilename: WideString; Page: Integer; X: Integer; Y: Integer;
                         Z: Integer; Layer: Integer): Integer; dispid 3;
    function GetSymbolCount(PageNo: Integer): Integer; dispid 46;
    function GetSymbolHandle(PageNo: Integer; SymbolIndex: Integer): Integer; dispid 47;
    procedure RemoveSymbol(PageNo: Integer; SymbolHandle: Integer); dispid 4;
    procedure SetSymbolText(PageNo: Integer; SymbolHandle: Integer; const Name: WideString;
                            const TypeText: WideString; const Article: WideString;
                            const FunctionText: WideString); dispid 5;
    procedure SetSymbolType(PageNo: Integer; SymbolHandle: Integer; AType: Integer); dispid 39;
    function GetSymbolType(PageNo: Integer; SymbolHandle: Integer; out AType: Integer): WordBool; dispid 61;
    procedure GetSymbolText(PageNo: Integer; SymbolHandle: Integer; out Name: WideString;
                            out TypeText: WideString; out Article: WideString;
                            out FunctionText: WideString); dispid 6;
    procedure GetSymbolBox(PageNo: Integer; SymbolHandle: Integer; BoxType: Integer;
                           out X1: Integer; out Y1: Integer; out X2: Integer; out Y2: Integer); dispid 7;
    procedure MoveSymbol(PageNo: Integer; SymbolHandle: Integer; NewPage: Integer; X: Integer;
                         Y: Integer; Z: Integer; Layer: Integer); dispid 8;
    procedure GetSymbolPos(PageNo: Integer; SymbolHandle: Integer; out X: Integer; out Y: Integer;
                           out Z: Integer; out Layer: Integer); dispid 9;
    procedure GetSymbolConnectionPosRel(PageNo: Integer; SymbolHandle: Integer;
                                        ConnectionIndex: Integer; out X: Integer; out Y: Integer;
                                        out Z: Integer); dispid 10;
    procedure GetSymbolConnectionIOstatus(PageNo: Integer; SymbolHandle: Integer;
                                          ConnectionIndex: Integer; out AStatus: Integer); dispid 43;
    function PlaceLine(PageNo: Integer; Layer: Integer; StartX: Integer; StartY: Integer;
                       StartZ: Integer; EndX: Integer; EndY: Integer; EndZ: Integer): Integer; dispid 11;
    function PlacePolyLine(PageNo: Integer; Layer: Integer; AList: OleVariant): Integer; dispid 34;
    function GetLineCount(PageNo: Integer): Integer; dispid 44;
    function GetLineHandle(PageNo: Integer; LineIndex: Integer): Integer; dispid 45;
    procedure SetLineOptions(Color: Integer; Style: Integer; Width: Integer); dispid 12;
    procedure UpdateDrawing; dispid 13;
    function GetConnections(PageIndex: Integer; SymbolHandle: Integer): OleVariant; dispid 14;
    function PlaceDynamicCable(PageNo: Integer; Layer: Integer; Direction: Integer;
                               AList: OleVariant): Integer; dispid 16;
    function GetSymbolConnectionCount(PageNo: Integer; SymbolHandle: Integer): Integer; dispid 17;
    procedure CreateJoins(PageNo: Integer); dispid 18;
    function GetSymbolDataFieldCount(PageNo: Integer; SymbolHandle: Integer): Integer; dispid 19;
    procedure GetSymbolDataField(PageNo: Integer; SymbolHandle: Integer; FieldIndex: Integer;
                                 out AName: WideString; out AText: WideString); dispid 20;
    procedure SetSymbolDataField(PageNo: Integer; SymbolHandle: Integer; FieldIndex: Integer;
                                 const AText: WideString); dispid 21;
    procedure SetSymbolConnectionText(PageNo: Integer; SymbolHandle: Integer;
                                      ConnectionIndex: Integer; const Name: WideString;
                                      const FunctionText: WideString; const LabelText: WideString;
                                      const Description: WideString); dispid 22;
    procedure GetSymbolConnectionText(PageNo: Integer; SymbolHandle: Integer;
                                      ConnectionIndex: Integer; out Name: WideString;
                                      out FunctionText: WideString; out LabelText: WideString;
                                      out Description: WideString); dispid 23;
    function PlaceHeader(PageIndex: Integer; const HeaderFilename: WideString): WordBool; dispid 24;
    function PlaceDynamicSymbol(const SymbolFilename: WideString; PageNo: Integer; X: Integer;
                                Y: Integer; Z: Integer; Layer: Integer; AList: OleVariant): Integer; dispid 25;
    procedure GetSymbolFileName(PageNo: Integer; SymbolHandle: Integer; out AFilename: WideString); dispid 71;
    procedure GetSymbolTitle(PageNo: Integer; SymbolHandle: Integer; out ATitle: WideString); dispid 26;
    procedure SetSymbolNameData(PageNo: Integer; SymbolHandle: Integer; AWidth: Integer;
                                AHeight: Integer; AOrigin: Integer; AOutLine: Integer;
                                AColor: Integer; ABold: Integer; AVisible: Integer; AAngle: Integer); dispid 27;
    procedure SetSymbolNamePos(PageNo: Integer; SymbolHandle: Integer; X: Integer; Y: Integer;
                               Z: Integer); dispid 28;
    procedure SetSymbolTypeTextData(PageNo: Integer; SymbolHandle: Integer; AWidth: Integer;
                                    AHeight: Integer; AOrigin: Integer; AOutLine: Integer;
                                    AColor: Integer; ABold: Integer; AVisible: Integer;
                                    AAngle: Integer); dispid 35;
    procedure SetSymbolTypeTextPos(PageNo: Integer; SymbolHandle: Integer; X: Integer; Y: Integer;
                                   Z: Integer); dispid 36;
    procedure SetSymbolArticleData(PageNo: Integer; SymbolHandle: Integer; AWidth: Integer;
                                   AHeight: Integer; AOrigin: Integer; AOutLine: Integer;
                                   AColor: Integer; ABold: Integer; AVisible: Integer;
                                   AAngle: Integer); dispid 37;
    procedure SetSymbolArticlePos(PageNo: Integer; SymbolHandle: Integer; X: Integer; Y: Integer;
                                  Z: Integer); dispid 38;
    procedure SetSymbolFunctionData(PageNo: Integer; SymbolHandle: Integer; AWidth: Integer;
                                    AHeight: Integer; AOrigin: Integer; AOutLine: Integer;
                                    AColor: Integer; ABold: Integer; AVisible: Integer;
                                    AAngle: Integer); dispid 29;
    procedure SetSymbolFunctionPos(PageNo: Integer; SymbolHandle: Integer; X: Integer; Y: Integer;
                                   Z: Integer); dispid 30;
    procedure GetPageHeaders(out Headers: OleVariant); dispid 31;
    procedure GetSymbolFileNames(out FileNames: OleVariant); dispid 32;
    procedure GetTextWidth(ACharWidth: Integer; const AText: WideString; out ATextWidth: Integer); dispid 33;
    function AddSpeedButton(const ABarName: WideString; const AHint: WideString;
                            const Caption_BMP: WideString; WinHandle: Integer; CmdID: Integer): Integer; dispid 66;
    function AddMenuItem(const AOwnerMenu: WideString; const ACaption: WideString;
                         WinHandle: Integer; CmdID: Integer): Integer; dispid 67;
    function RemoveCmdItem(AItem: Integer): WordBool; dispid 68;
    function EnableCmdItem(AItem: Integer; AVisible: WordBool; AEnabled: WordBool): WordBool; dispid 69;
    procedure CreateTermList; dispid 62;
    procedure FreeTermList; dispid 63;
    function GetTermListCount: Integer; dispid 64;
    procedure GetTermList(ListIndex: Integer; out AList: OleVariant); dispid 65;
    procedure BringToFront; dispid 70;
    function GetProgramState: Integer; dispid 72;
    function AddMenu(const AOwnerMenu: WideString; const ACaption: WideString;
                     out AMenuName: WideString): Integer; dispid 74;
    procedure SetPageReference(PageNo: Integer; AFlags: Integer); dispid 75;
    procedure SetSymbolConnTextsVisible(PageNo: Integer; SymbolHandle: Integer; ATexts: Integer;
                                        AVisible: Integer); dispid 76;
    function PlaceText(const AText: WideString; Page: Integer; X: Integer; Y: Integer; Z: Integer;
                       Layer: Integer): Integer; dispid 77;
    procedure SetLayVisible(Layer: Integer; AVisible: Integer); dispid 78;
    procedure SymbolMenu(const ALibrary: WideString); dispid 79;
    procedure DontShutDownOnRelease; dispid 80;
    function GetProgramVersion: WideString; dispid 81;
    function GetProgramEnv(const EnvField: WideString): WideString; dispid 82;
    procedure UpdateMenus; dispid 83;
    procedure RotateSymbol(PageNo: Integer; SymbolHandle: Integer; Angle: Integer); dispid 84;
    procedure SetProjectTextsVisible(ATexts: Integer; AVisible: Integer); dispid 85;
    procedure TransferLibType(SrcProjectNo: Integer; DstProjectNo: Integer;
                              const AFilename: WideString); dispid 86;
    function GetActiveProjectHandle: Integer; dispid 87;
    function GetFunctionID(const FunctionName: WideString): Integer; dispid 88;
    procedure FunctionActivate(AID: Integer; Activate: WordBool); dispid 89;
    procedure FunctionVisible(AID: Integer; Visible: WordBool); dispid 90;
    procedure FunctionEnable(AID: Integer; Enable: WordBool); dispid 91;
    procedure SetFunctionMsg(AID: Integer; WinHandle: Integer; Msg: Integer); dispid 92;
    procedure GetProjectName(out AFilename: WideString); dispid 93;
    function GetActiveProject: Integer; dispid 94;
    procedure MoveLay(PageNo: Integer; Layer: Integer; NewLayer: Integer); dispid 95;
    procedure RemoveLay(PageNo: Integer; Layer: Integer); dispid 96;
    procedure SendOnTerminate(WinHandle: Integer; Msg: Integer); dispid 97;
    procedure CancelSaveCheck(Cancel: WordBool); dispid 98;
    procedure LockUserInterface(Happ: Integer; Lock: WordBool); dispid 99;
    procedure SetTextData(PageNo: Integer; TextHandle: Integer; AWidth: Integer; AHeight: Integer;
                          AOrigin: Integer; AOutLine: Integer; AColor: Integer; ABold: Integer;
                          AVisible: Integer; AAngle: Integer); dispid 100;
    procedure SetListOptions(ListType: Integer; PageNo: Integer; PrimSort: Integer;
                             SecSort: Integer; Options: Integer); dispid 101;
    procedure GetApplicationData(out Data: OleVariant); dispid 102;
    procedure SetApplicationData(Data: OleVariant); dispid 103;
    procedure GetProjectTitle(out ATitle: WideString); dispid 104;
    function GetMainWindowHandle: Integer; dispid 105;
    procedure NameGetSymbolHandles(APageType: Integer; AObjectType: Integer;
                                   const AName: WideString; out Handles: OleVariant); dispid 106;
    procedure SetSymbolConnPinNameData(PageNo: Integer; SymbolHandle: Integer;
                                       ConnectionIndex: Integer; AWidth: Integer; AHeight: Integer;
                                       AOrigin: Integer; AOutLine: Integer; AColor: Integer;
                                       ABold: Integer; AVisible: Integer; AAngle: Integer); dispid 107;
    procedure JoinParkedLines(PageNo: Integer); dispid 108;
    procedure SetSymbolScale(PageNo: Integer; SymbolHandle: Integer; AScale: Double); dispid 109;
    procedure RemovePage(PageNo: Integer); dispid 110;
    procedure GetAreas(PageNo: Integer; Layer: Integer; APenIndex: Integer; AMode: Integer;
                       out AList: OleVariant); dispid 111;
    function PlacePicture(const AFilename: WideString; Page: Integer; X: Integer; Y: Integer;
                          Z: Integer; Width: Integer; Height: Integer): Integer; dispid 112;
    procedure RemovePicture(PageNo: Integer; PictureHandle: Integer); dispid 113;
    function GetPictureCount(PageNo: Integer): Integer; dispid 114;
    function GetPictureHandle(PageNo: Integer; PictureIndex: Integer): Integer; dispid 115;
    procedure GetHeaderFileName(PageNo: Integer; out AFilename: WideString); dispid 116;
    procedure GetPageParms(PageNo: Integer; out ASize: Integer; out AType: Integer;
                           out AFunction: Integer; out AScale: Double); dispid 117;
    function GetSymbolDataFields(PageNo: Integer; SymbolHandle: Integer; out Fields: OleVariant;
                                 out Data: OleVariant): WordBool; dispid 118;
    function SetSymbolDataFields(PageNo: Integer; SymbolHandle: Integer; Data: OleVariant): WordBool; dispid 119;
    function CreateCableList(SortIndex: Integer; Lay: Integer; const StartPage: WideString;
                             const EndPage: WideString; Options: Integer): WordBool; dispid 120;
    procedure FreeCableList; dispid 121;
    function GetCableListCount: Integer; dispid 122;
    procedure GetCableList(ListIndex: Integer; out AList: OleVariant); dispid 123;
    procedure GetTermListData(ListIndex: Integer; DataType: Integer; out Symbol: Integer;
                              out Connection: Integer); dispid 124;
    procedure GetCableListData(ListIndex: Integer; DataType: Integer; out Symbol: Integer;
                               out Connection: Integer); dispid 125;
    procedure ActivateExtPopup(WinHandle: Integer; Msg: Integer); dispid 126;
    procedure ShowPopupMenu(ExtMenu: WordBool); dispid 127;
    procedure GetSelectionStatus(out AStatus: Integer); dispid 128;
    procedure GetSelection(AType: Integer; out AList: OleVariant); dispid 129;
    procedure SetPageTitle(PageNo: Integer; const ATitle: WideString); dispid 130;
    procedure ShowPickMenu(ShowIt: WordBool); dispid 131;
    procedure GetPopupMenuPos(InPixels: WordBool; out X: Integer; out Y: Integer); dispid 132;
    procedure SetPageExclIndexList(PageNo: Integer; Excl: WordBool); dispid 133;
    procedure SetElectricalLines(Electr: WordBool); dispid 134;
    function GetCurrentPageIndex: Integer; dispid 135;
    procedure SetPageType(PageNo: Integer; AType: Integer); dispid 136;
    procedure ClearSelection; dispid 137;
    procedure CreateNet(PageNo: Integer; SymbolHandle: Integer; ConnevtionIndex: Integer); dispid 138;
    procedure GetSelectionArea(out Left: Integer; out Right: Integer; out Bottom: Integer;
                               out Top: Integer; out AreaSelect: WordBool); dispid 139;
    procedure CreateCompList; dispid 140;
    procedure FreeCompList; dispid 141;
    function GetCompListCount: Integer; dispid 142;
    procedure CreatePartList; dispid 143;
    procedure FreePartList; dispid 144;
    function GetPartListCount: Integer; dispid 145;
    function GetDataField(const AName: WideString; ADataNo: Integer; ADataStatus: Integer;
                          AIndex: Integer): WideString; dispid 146;
    procedure ConnectSymbol(PageNo: Integer; SymbolHandle: Integer); dispid 147;
    function PlaceCableWire(PageNo: Integer; CableHandle: Integer; DownRight: WordBool; X: Integer;
                            Y: Integer; Z: Integer): Integer; dispid 148;
    procedure RemoveLine(PageNo: Integer; LineHandle: Integer); dispid 149;
    function GetLinePointCount(PageNo: Integer; LineHandle: Integer): Integer; dispid 150;
    function GetLinePointJoinIndex(PageNo: Integer; LineHandle: Integer; PointIndex: Integer): Integer; dispid 151;
    function GetSymbolConnectionJoinIndex(PageNo: Integer; SymbolHandle: Integer;
                                          ConnectionIndex: Integer): Integer; dispid 152;
    procedure ChangeLay(Layer: Integer); dispid 153;
    procedure ShowMsg(const AMsg: WideString; const BtnText: WideString; AType: Integer;
                      WinHandle: Integer; WinMsg: Integer); dispid 154;
    procedure CloseMsg; dispid 155;
    function InsertMenuItem(const AOwnerMenu: WideString; const ACaption: WideString;
                            AIndex: Integer; WinHandle: Integer; CmdID: Integer): Integer; dispid 156;
    function GetSymbolConnectionJoined(PageNo: Integer; SymbolHandle: Integer;
                                       ConnectionIndex: Integer): WordBool; dispid 157;
    procedure FindFirstNetInputOutput(AStatus: Integer; PageNo: Integer; SymbolHandle: Integer;
                                      ConnectionIndex: Integer; out PgNo: Integer;
                                      out SymHandle: Integer; out ConIndex: Integer); dispid 158;
    procedure FindFirstNetSymbolType(AType: Integer; PageNo: Integer; SymbolHandle: Integer;
                                     ConnectionIndex: Integer; out PgNo: Integer;
                                     out SymHandle: Integer; out ConIndex: Integer); dispid 159;
    function DataBaseMenu(const Articles: WideString; out AArticle: WideString;
                          out AType: WideString; out AFunction: WideString): WordBool; dispid 160;
    procedure SendOnSelection(WinHandle: Integer; Msg: Integer); dispid 161;
    function HighLightNet(PageNo: Integer; SymbolHandle: Integer; ConnectionIndex: Integer;
                          AColor: Integer): WordBool; dispid 162;
    function HighLightSymbol(PageNo: Integer; SymbolHandle: Integer; AColor: Integer): WordBool; dispid 163;
    function FileNotSaved(PageNo: Integer): WordBool; dispid 164;
    function SetSymbolQuantity(PageNo: Integer; SymbolHandle: Integer; Quantity: Double): WordBool; dispid 165;
    function SetSymbolConnTextData(PageNo: Integer; SymbolHandle: Integer;
                                   ConnectionIndex: Integer; TextIndex: Integer; AWidth: Integer;
                                   AHeight: Integer; AOrigin: Integer; AOutLine: Integer;
                                   AColor: Integer; ABold: Integer; AVisible: Integer;
                                   AAngle: Integer): WordBool; dispid 166;
    function GetNetListExt(NetIndex: Integer; Flags: Integer; out AList: OleVariant): WordBool; dispid 167;
    function CopySymbol(SrcProjectNo: Integer; DstProjectNo: Integer; SrcPageNo: Integer;
                        DstPageNo: Integer; SrcSymbolHandle: Integer; X: Integer; Y: Integer;
                        Z: Integer; Layer: Integer): Integer; dispid 168;
    function PinSwop(PageNo: Integer; SymbolHandle: Integer; Connection1Index: Integer;
                     Connection2Index: Integer): WordBool; dispid 169;
    function CopyPage(SrcProjectNo: Integer; DstProjectNo: Integer; SrcPageNo: Integer;
                      DstPageNo: Integer): WordBool; dispid 170;
    function CopyEmptyPage(SrcProjectNo: Integer; DstProjectNo: Integer; SrcPageNo: Integer;
                           DstPageNo: Integer): WordBool; dispid 171;
    function SetOleProject(ProjectNo: Integer): WordBool; dispid 172;
    function InsertProject(PageNo: Integer; const AFilename: WideString;
                           const APageNumber: WideString; Options: Integer): WordBool; dispid 173;
    function SetPageNumber(PageNo: Integer; const ANumber: WideString): WordBool; dispid 174;
    function GetPageNumber(PageNo: Integer; out ANumber: WideString): WordBool; dispid 175;
    function SetPageTagData(PageNo: Integer; AGroupNo: Integer; AList: OleVariant): WordBool; dispid 176;
    function GetPageTagData(PageNo: Integer; AGroupNo: Integer; out AList: OleVariant): WordBool; dispid 177;
    function GetPageTagNames(PageNo: Integer; AGroupNo: Integer; out AList: OleVariant): WordBool; dispid 178;
    function CreatePageTags(PageNo: Integer): WordBool; dispid 179;
    function GetInsertionProints(PageNo: Integer; MinDist: Integer; out X1: Integer;
                                 out Y1: Integer; out Z1: Integer; out X2: Integer;
                                 out Y2: Integer; out Z2: Integer): Integer; dispid 181;
    function LoadSubDrawing(const AFilename: WideString): WordBool; dispid 182;
    function ReleaseSubDrawing: WordBool; dispid 183;
    function PlaceSubDrawing(PageNo: Integer; X: Integer; Y: Integer; Z: Integer; Layer: Integer;
                             Group: Integer; Options: Integer): WordBool; dispid 184;
    function GetPageData(PageNo: Integer; out Defs: OleVariant; out Data: OleVariant): WordBool; dispid 180;
    function SetPageData(PageNo: Integer; Data: OleVariant): WordBool; dispid 185;
    function MirrorSymbol(PageNo: Integer; SymbolHandle: Integer; Vertical: WordBool;
                          FromCurDir: WordBool): WordBool; dispid 186;
    function OpenProjectTemplate(const AFilename: WideString): Integer; dispid 187;
    function GetTextCount(PageNo: Integer): Integer; dispid 188;
    function GetTextHandle(PageNo: Integer; TextIndex: Integer): Integer; dispid 189;
    function GetText(PageNo: Integer; TextHandle: Integer; out AText: WideString): WordBool; dispid 190;
    function SetText(PageNo: Integer; TextHandle: Integer; const AText: WideString): WordBool; dispid 191;
    function CreateConnList(ListType: Integer; Layer: Integer): WordBool; dispid 192;
    function FreeConnList: WordBool; dispid 193;
    function GetConnListCount: Integer; dispid 194;
    function GetConnList(ConnIndex: Integer; Flags: Integer; out AList: OleVariant): WordBool; dispid 195;
    procedure SendOnProject(WndHandle: Integer; Msg: Integer); dispid 196;
    function SetCursor(ACursor: Integer): Integer; dispid 197;
    function ApplicationDataReadSection(PageNo: Integer; const Section: WideString;
                                        out Sata: OleVariant): WordBool; dispid 198;
    function ApplicationDataReadString(PageNo: Integer; const Section: WideString;
                                       const Ident: WideString; const Default: WideString;
                                       out Value: WideString): WordBool; dispid 199;
    function ApplicationDataWriteSection(PageNo: Integer; const Section: WideString;
                                         Data: OleVariant): WordBool; dispid 200;
    function ApplicationDataWriteString(PageNo: Integer; const Section: WideString;
                                        const Ident: WideString; const Value: WideString): WordBool; dispid 201;
    function GetReferenceID(PageNo: Integer; SymbolHandle: Integer; out AFunction: WideString;
                            out ALocation: WideString; out AName: WideString;
                            out ASubName: WideString): WordBool; dispid 202;
    function SetReferenceID(PageNo: Integer; SymbolHandle: Integer; const AFunction: WideString;
                            const ALocation: WideString; const AName: WideString;
                            const ASubName: WideString): WordBool; dispid 203;
    function CreateReferenceID(AType: Integer; const AID: WideString; const ADiscr: WideString): WordBool; dispid 204;
    function UpdateProjectFromDB(PageNo: Integer; SymbolHandle: Integer): WordBool; dispid 205;
    function GetSymbolCompGroup(PageNo: Integer; SymbolHandle: Integer; out AGroupNo: Integer;
                                out APosNo: Integer): WordBool; dispid 206;
    function CreateSymbolDataField(const AName: WideString): WordBool; dispid 207;
    function GetPageUsedDataFields(AType: Integer; PageNo: Integer; out Fields: OleVariant): WordBool; dispid 208;
    function PlaceCableScreen(PageNo: Integer; CableHandle: Integer): Integer; dispid 209;
    function GetPageTemplateNames(PageNo: Integer; out AList: OleVariant): WordBool; dispid 210;
    function GetSymbolStates(PageNo: Integer; SymbolHandle: Integer; out States: OleVariant): WordBool; dispid 211;
    function SetSymbolState(PageNo: Integer; SymbolHandle: Integer; AState: Integer): WordBool; dispid 212;
    function GetSymbolState(PageNo: Integer; SymbolHandle: Integer; out AState: Integer): WordBool; dispid 213;
    function Print(PageNo: Integer): WordBool; dispid 214;
    function LoadMechPageCompList(PageNo: Integer): WordBool; dispid 215;
    procedure SendOnDesignSymbol(WndHandle: Integer; Msg: Integer); dispid 216;
    function SortCreatedList(AListType: Integer; APrimSort: Integer; ASecSort: Integer): WordBool; dispid 217;
    function SetObjectVisible(APageNo: Integer; ObjectType: Integer; ObjectHandle: Integer;
                              Visible: WordBool): WordBool; dispid 218;
    function SetProjectTitle(const ATitle: WideString): WordBool; dispid 219;
    function SetSymbolConnectionPolygon(PageNo: Integer; SymbolHandle: Integer;
                                        ConnectionIndex: Integer; AShape: Integer;
                                        BordorColor: Integer; FillColor: Integer;
                                        const Coors: OleVariant; Options: Integer): WordBool; dispid 220;
    function EXCreateCompList(const ParamStr: WideString): WordBool; dispid 221;
    function EXCreatePartList(const ParamStr: WideString): WordBool; dispid 222;
    function SymbolApplDataReadString(PageNo: Integer; SymbolHandle: Integer;
                                      const Section: WideString; const Ident: WideString;
                                      const Default: WideString; out Value: WideString): WordBool; dispid 223;
    function SymbolApplDataWriteString(PageNo: Integer; SymbolHandle: Integer;
                                       const Section: WideString; const Ident: WideString;
                                       const Value: WideString): WordBool; dispid 224;
    function GetDBFieldValue(const SearchField: WideString; const SearchValue: WideString;
                             const FieldName: WideString; out Value: WideString): WordBool; dispid 225;
  end;

// *********************************************************************//
// DispIntf:  IPCsApplicationEvents
// Flags:     (0)
// GUID:      {4576D15E-0EE2-4B1B-8AC3-2E820FEE1222}
// *********************************************************************//
  IPCsApplicationEvents = dispinterface
    ['{4576D15E-0EE2-4B1B-8AC3-2E820FEE1222}']
    function QuitEvent: HResult; dispid 201;
  end;

// *********************************************************************//
// Interface: IPCsApplication2
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {F6256177-1C51-460F-84C7-AB1905D508E5}
// *********************************************************************//
  IPCsApplication2 = interface(IPCsApplication)
    ['{F6256177-1C51-460F-84C7-AB1905D508E5}']
    function GetDocumentFromHandle(Handle: Integer): IPCsDocument; safecall;
    function CreateObject(ClassID: TGUID; const Owner: IUnknown): IUnknown; safecall;
    function Equals(const Obj: IPCsBase): WordBool; safecall;
    function Get_Application: IPCsApplication2; safecall;
    property Application: IPCsApplication2 read Get_Application;
  end;

// *********************************************************************//
// DispIntf:  IPCsApplication2Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {F6256177-1C51-460F-84C7-AB1905D508E5}
// *********************************************************************//
  IPCsApplication2Disp = dispinterface
    ['{F6256177-1C51-460F-84C7-AB1905D508E5}']
    function GetDocumentFromHandle(Handle: Integer): IPCsDocument; dispid 301;
    function CreateObject(ClassID: {NOT_OLEAUTO(TGUID)}OleVariant; const Owner: IUnknown): IUnknown; dispid 302;
    function Equals(const Obj: IPCsBase): WordBool; dispid 303;
    property Application: IPCsApplication2 readonly dispid 304;
    property Documents: IPCsDocuments readonly dispid 400;
    property AvailableFonts: OleVariant readonly dispid 402;
    property PenColors[PenColor: PsPenColor]: OLE_COLOR readonly dispid 403;
    property Preferences: IPCsApplicationPreferences readonly dispid 404;
    property SymbolLibraryMenu: IPCsSymbolMenu readonly dispid 405;
    property StatusBar: WideString dispid 406;
    property Visible: WordBool dispid 408;
    procedure Quit; dispid 409;
    procedure DontShutDownOnLastRelease; dispid 410;
    function CreateVirtualPage: IPCsPage; dispid 411;
    property ActiveComponentDatabase: IPCsComponentDatabase readonly dispid 412;
    function ExecuteModule(const Filename: WideString; const ParameterString: WideString): WordBool; dispid 413;
    procedure Redraw; dispid 414;
    property ActiveDocument: IPCsDocument dispid 201;
    procedure ZoomAll; dispid 202;
    function LoadString(Identifier: Integer): WideString; dispid 203;
    function SetOleStatusbarText(PanelIndex: Integer; const Text: WideString; Width: Integer;
                                 Alignment: Shortint): TWindowHandle; dispid 204;
    property Dialogs[Kind: psDialogKind]: IPCsDialog readonly dispid 205;
    procedure Activate; dispid 206;
    property ExplorerWindow: IPCsExplorerWindow readonly dispid 207;
    function CreateNewLibType: IPCsLibType; dispid 208;
  end;

// *********************************************************************//
// DispIntf:  IPCsDocumentEvents
// Flags:     (0)
// GUID:      {0061C56D-4338-4314-99DD-FB039E7BB236}
// *********************************************************************//
  IPCsDocumentEvents = dispinterface
    ['{0061C56D-4338-4314-99DD-FB039E7BB236}']
    function Close: HResult; dispid 201;
    function CloseQuery(CanClose: WordBool): HResult; dispid 202;
  end;

// *********************************************************************//
// Interface: IPCsDocument2
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {43F030B6-1F4C-47CE-9A07-98BC1ECFAE2F}
// *********************************************************************//
  IPCsDocument2 = interface(IPCsDocument)
    ['{43F030B6-1F4C-47CE-9A07-98BC1ECFAE2F}']
    function Copy: IPCsDocument2; safecall;
    function Equals(const Obj: IPCsBase): WordBool; safecall;
    function IsErased: WordBool; safecall;
  end;

// *********************************************************************//
// DispIntf:  IPCsDocument2Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {43F030B6-1F4C-47CE-9A07-98BC1ECFAE2F}
// *********************************************************************//
  IPCsDocument2Disp = dispinterface
    ['{43F030B6-1F4C-47CE-9A07-98BC1ECFAE2F}']
    function Copy: IPCsDocument2; dispid 226;
    function Equals(const Obj: IPCsBase): WordBool; dispid 301;
    function IsErased: WordBool; dispid 302;
    property Drawing: IPCsDrawing readonly dispid 201;
    property Application: IPCsApplication readonly dispid 202;
    procedure Close(SaveChanges: WordBool; const Filename: WideString); dispid 203;
    property FullName: WideString readonly dispid 206;
    function ChangeFilename(const Filename: WideString): WordBool; dispid 207;
    property CharSetName: WideString dispid 208;
    procedure Activate; dispid 210;
    property WindowHandle: TWindowHandle readonly dispid 204;
    property IsNewFile: WordBool readonly dispid 211;
    procedure ZoomAll; dispid 212;
    property ActivePage: IPCsPage dispid 205;
    property Saved: WordBool dispid 209;
    property DataLink: IPCsDataLink readonly dispid 213;
    procedure Save; dispid 214;
    procedure SaveAs(const Filename: WideString); dispid 215;
    procedure Print; dispid 216;
    property SelectStatus: IPCsSelectStatus readonly dispid 217;
    property Handle: LongWord readonly dispid 218;
    procedure BeginUndo(const Title: WideString); dispid 219;
    procedure EndUndo; dispid 220;
    procedure Undo; dispid 221;
    procedure Select(const Page: IPCsPage; const Layer: IPCsLayer; const FilterString: WideString); dispid 222;
    procedure SelectInWindow(AreaLo: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant;
                             AreaHi: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Page: IPCsPage;
                             const Layer: IPCsLayer; const FilterString: WideString); dispid 223;
    function GotoObject(const AObject: IUnknown; Options: psGotoObjectOption): WordBool; dispid 224;
  end;

// *********************************************************************//
// DispIntf:  IPCsApplicationPreferencesEvents
// Flags:     (0)
// GUID:      {836113FE-0903-4C9D-B517-33BFBD5B7191}
// *********************************************************************//
  IPCsApplicationPreferencesEvents = dispinterface
    ['{836113FE-0903-4C9D-B517-33BFBD5B7191}']
  end;

// *********************************************************************//
// DispIntf:  IPCsStatusPointEvents
// Flags:     (0)
// GUID:      {2DA70029-2E21-4962-A583-A4034CD9D819}
// *********************************************************************//
  IPCsStatusPointEvents = dispinterface
    ['{2DA70029-2E21-4962-A583-A4034CD9D819}']
  end;

// *********************************************************************//
// Interface: IPCsLibType2
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {7203263B-5703-4251-B3A4-827118605237}
// *********************************************************************//
  IPCsLibType2 = interface(IPCsLibType)
    ['{7203263B-5703-4251-B3A4-827118605237}']
    function GraphicsCompare(Options: Integer): WordBool; safecall;
    function Equals(const Obj: IPCsBase): WordBool; safecall;
    procedure AssignTo(const Obj: IPCsBase); safecall;
    function Get_Application: IPCsApplication2; safecall;
    property Application: IPCsApplication2 read Get_Application;
  end;

// *********************************************************************//
// DispIntf:  IPCsLibType2Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {7203263B-5703-4251-B3A4-827118605237}
// *********************************************************************//
  IPCsLibType2Disp = dispinterface
    ['{7203263B-5703-4251-B3A4-827118605237}']
    function GraphicsCompare(Options: Integer): WordBool; dispid 401;
    function Equals(const Obj: IPCsBase): WordBool; dispid 402;
    procedure AssignTo(const Obj: IPCsBase); dispid 403;
    property Application: IPCsApplication2 readonly dispid 404;
    property SymbolType: psSymbolType dispid 302;
    property Name: WideString dispid 304;
    property FileTime: Integer dispid 305;
    property IsReferenceOnly: WordBool readonly dispid 306;
    property Handle: LongWord readonly dispid 307;
    property Parent: IPCsLibTypes readonly dispid 301;
    function GetRect(AType: PsSymbolRectType): {NOT_OLEAUTO(TPCsRect)}OleVariant; dispid 309;
    property IsElectrical: WordBool dispid 310;
    property IsMechanical: WordBool dispid 311;
    procedure Delete; dispid 312;
    property IsUsed: WordBool readonly dispid 313;
    property ObjectType: PsPCsObjectType readonly dispid 450;
    property Lines: IPCsLines readonly dispid 201;
    property Connections: IPCsLibTypeConnections readonly dispid 202;
    property Arcs: IPCsArcs readonly dispid 203;
    property Texts: IPCsTexts readonly dispid 204;
    property Datafields: IPCsDatafields readonly dispid 205;
    property Title: WideString dispid 303;
    property TitleLC[LCID: Integer]: WideString dispid 206;
    property States: OleVariant readonly dispid 207;
    property SymbolTexts: IPCsSymbolTexts readonly dispid 208;
  end;

// *********************************************************************//
// DispIntf:  IPCsDrawingEvents
// Flags:     (0)
// GUID:      {E069BDF3-AE9A-43A6-902F-6D8CF814E1D4}
// *********************************************************************//
  IPCsDrawingEvents = dispinterface
    ['{E069BDF3-AE9A-43A6-902F-6D8CF814E1D4}']
  end;

// *********************************************************************//
// Interface: IPCsDrawing2
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {9CD75DDE-BDF9-446F-BB19-3BEDF942A0A1}
// *********************************************************************//
  IPCsDrawing2 = interface(IPCsDrawing)
    ['{9CD75DDE-BDF9-446F-BB19-3BEDF942A0A1}']
    function GetUsedLibtypes: OleVariant; safecall;
    function Equals(const Obj: IPCsBase): WordBool; safecall;
    function IsErased: WordBool; safecall;
    function Get_Logos(Index: Integer): IPCsExternalObject; safecall;
    procedure Set_Logos(Index: Integer; const Value: IPCsExternalObject); safecall;
    function LoadLogo(Index: Integer; const Filename: WideString): WordBool; safecall;
    function Get_LogosCount: Integer; safecall;
    function ReplaceSymbol(const SymbolToReplace: IPCsSymbol;
                           const SymbolToReplaceWith: IPCsSymbol; Options: Integer): IPCsSymbol; safecall;
    property Logos[Index: Integer]: IPCsExternalObject read Get_Logos write Set_Logos;
    property LogosCount: Integer read Get_LogosCount;
  end;

// *********************************************************************//
// DispIntf:  IPCsDrawing2Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {9CD75DDE-BDF9-446F-BB19-3BEDF942A0A1}
// *********************************************************************//
  IPCsDrawing2Disp = dispinterface
    ['{9CD75DDE-BDF9-446F-BB19-3BEDF942A0A1}']
    function GetUsedLibtypes: OleVariant; dispid 301;
    function Equals(const Obj: IPCsBase): WordBool; dispid 302;
    function IsErased: WordBool; dispid 303;
    property Logos[Index: Integer]: IPCsExternalObject dispid 304;
    function LoadLogo(Index: Integer; const Filename: WideString): WordBool; dispid 305;
    property LogosCount: Integer readonly dispid 306;
    function ReplaceSymbol(const SymbolToReplace: IPCsSymbol;
                           const SymbolToReplaceWith: IPCsSymbol; Options: Integer): IPCsSymbol; dispid 307;
    property Pages: IPCsPages readonly dispid 201;
    property LibTypes: IPCsLibTypes readonly dispid 202;
    property Layers: IPCsLayers readonly dispid 205;
    property Application: IPCsApplication readonly dispid 206;
    property Parent: IPCsDocument readonly dispid 207;
    property Title: WideString dispid 203;
    function GetSymbolFromHandle(PageNo: Integer; SymbolHandle: Integer): IPCsSymbol; dispid 204;
    property ActiveLayer: IPCsLayer dispid 208;
    property ProjectData[Index: Integer]: WideString dispid 209;
    property ProjectDataDefs: OleVariant dispid 210;
    property PageDataDefs: OleVariant dispid 211;
    property SymbolDataDefs: OleVariant dispid 214;
    property ReferenceDesignations: IPCsReferenceDesignations readonly dispid 212;
    function SetProjectData(const FieldName: WideString; const Value: WideString): Integer; dispid 213;
    property Lists: IPCsLists readonly dispid 215;
    procedure OptimizeLibtypes; dispid 216;
    function LoadSubDrawing(const Filename: WideString): IPCsSubDrawing; dispid 217;
    function GetLineFromHandle(PageNo: Integer; LineHandle: Integer): IPCsLine; dispid 218;
    function GetTextFromHandle(PageNo: Integer; TextHandle: Integer): IPCsText; dispid 219;
    property SystemVariable[const Variable: WideString]: OleVariant dispid 220;
    property Preferences: IPCsDrawingPreferences readonly dispid 221;
    procedure SynchronizePLCData(const ParamStr: WideString); dispid 222;
    procedure ApplyIntelligentInvisibility(const ParamStr: WideString); dispid 223;
    function UpdateFromDatabase(Reserved: Integer): WordBool; dispid 224;
    property Document: IPCsDocument readonly dispid 225;
  end;

// *********************************************************************//
// DispIntf:  IPCsLibTypesEvents
// Flags:     (0)
// GUID:      {DB192BAD-CAB3-45D6-AC9D-B92A33CADB48}
// *********************************************************************//
  IPCsLibTypesEvents = dispinterface
    ['{DB192BAD-CAB3-45D6-AC9D-B92A33CADB48}']
  end;

// *********************************************************************//
// DispIntf:  IPCsPagesEvents
// Flags:     (0)
// GUID:      {DD56960D-C565-48AE-812B-3B03D1D49E8B}
// *********************************************************************//
  IPCsPagesEvents = dispinterface
    ['{DD56960D-C565-48AE-812B-3B03D1D49E8B}']
  end;

// *********************************************************************//
// Interface: IPCsPages2
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {C4DE62DA-2FAA-4395-BA73-6B5C6BD16232}
// *********************************************************************//
  IPCsPages2 = interface(IPCsPages)
    ['{C4DE62DA-2FAA-4395-BA73-6B5C6BD16232}']
    procedure CopyPages(FromIndex: Integer; ToIndex: Integer; const DestDrawing: IPCsDrawing;
                        InsertAtIndex: Integer); safecall;
    procedure MovePages(FromIndex: Integer; ToIndex: Integer; MoveToIndex: Integer); safecall;
    function Get_Application: IPCsApplication2; safecall;
    function IsErased: WordBool; safecall;
    property Application: IPCsApplication2 read Get_Application;
  end;

// *********************************************************************//
// DispIntf:  IPCsPages2Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {C4DE62DA-2FAA-4395-BA73-6B5C6BD16232}
// *********************************************************************//
  IPCsPages2Disp = dispinterface
    ['{C4DE62DA-2FAA-4395-BA73-6B5C6BD16232}']
    procedure CopyPages(FromIndex: Integer; ToIndex: Integer; const DestDrawing: IPCsDrawing;
                        InsertAtIndex: Integer); dispid 302;
    procedure MovePages(FromIndex: Integer; ToIndex: Integer; MoveToIndex: Integer); dispid 303;
    property Application: IPCsApplication2 readonly dispid 301;
    function IsErased: WordBool; dispid 304;
    property Item[Index: Integer]: IPCsPage readonly dispid 0; default;
    property Count: Integer readonly dispid 203;
    property _NewEnum: IEnumVARIANT readonly dispid -4;
    function Add(PageFunction: psPageFunction; const Template: WideString): IPCsPage; dispid 204;
    function Insert(Index: Integer; PageFunction: psPageFunction; const Template: WideString): IPCsPage; dispid 201;
    procedure RemovePages(IndexFrom: Integer; IndexTo: Integer); dispid 202;
  end;

// *********************************************************************//
// Interface: IPCsSymbol2
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {669A39E0-9570-4494-BBB8-2AF3C5BF9249}
// *********************************************************************//
  IPCsSymbol2 = interface(IPCsSymbol)
    ['{669A39E0-9570-4494-BBB8-2AF3C5BF9249}']
    function Get_SymbolData(Index: Integer): WideString; safecall;
    procedure Set_SymbolData(Index: Integer; const Value: WideString); safecall;
    function Get_SymbolDataCount: Integer; safecall;
    function Equals(const Obj: IPCsBase): WordBool; safecall;
    function IsErased: WordBool; safecall;
    property SymbolData[Index: Integer]: WideString read Get_SymbolData write Set_SymbolData;
    property SymbolDataCount: Integer read Get_SymbolDataCount;
  end;

// *********************************************************************//
// DispIntf:  IPCsSymbol2Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {669A39E0-9570-4494-BBB8-2AF3C5BF9249}
// *********************************************************************//
  IPCsSymbol2Disp = dispinterface
    ['{669A39E0-9570-4494-BBB8-2AF3C5BF9249}']
    property SymbolData[Index: Integer]: WideString dispid 451;
    property SymbolDataCount: Integer readonly dispid 452;
    function Equals(const Obj: IPCsBase): WordBool; dispid 401;
    function IsErased: WordBool; dispid 403;
    property LibType: IPCsLibType readonly dispid 402;
    property IsVMirrored: WordBool dispid 405;
    property IsHMirrored: WordBool dispid 406;
    property SymbolTexts: IPCsSymbolTexts readonly dispid 407;
    property Page: IPCsPage readonly dispid 408;
    property Layer: IPCsLayer dispid 409;
    property FullName: WideString dispid 302;
    function GetRect(AType: PsSymbolRectType): {NOT_OLEAUTO(TPCsRect)}OleVariant; dispid 303;
    property Connections: IPCsConnections readonly dispid 304;
    property Rotation: Integer dispid 305;
    property ScaleFactor: Double dispid 306;
    function Explode(const ADestination: IPCsFigure): WordBool; dispid 307;
    property SymbolType1: psSymbolType dispid 301;
    property SymbolType2: psSymbolType dispid 308;
    procedure ApplicationDataWriteString(const ASection: WideString; const AIdent: WideString;
                                         const AValue: WideString); dispid 309;
    function ApplicationDataReadString(const ASection: WideString; const AIdent: WideString;
                                       const ADefault: WideString): WideString; dispid 310;
    property Datafields: IPCsDatafields readonly dispid 311;
    procedure Delete; dispid 312;
    property ObjectType: PsPCsObjectType readonly dispid 450;
    property ComponentGroupNo: Integer dispid 313;
    property ComponentPosNo: Integer dispid 314;
    property Reference: IPCsReference readonly dispid 315;
    property WithReference: WordBool dispid 316;
    function UpdateFromDatabase(Reserved: Integer): WordBool; dispid 317;
    property Visible: WordBool dispid 201;
    property Position: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant dispid 202;
    property LStatus: Integer dispid 203;
    function Move(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Figure: IPCsFigure): WordBool; dispid 204;
    procedure Rotate(DeltaAngle: Integer; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 205;
    procedure Scale(Factor: Single; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 208;
    property Handle: LongWord readonly dispid 206;
    procedure GetPosition(out X: Integer; out Y: Integer; out Z: Integer); dispid 207;
    procedure SetPosition(X: Integer; Y: Integer; Z: Integer); dispid 209;
    property State: Integer dispid 250;
    function ChangeDrawingOrder(AAction: psChangeDrawingOrder): WordBool; dispid 210;
  end;

// *********************************************************************//
// Interface: IPCsSymbolsEnumerable
// Flags:     (320) Dual OleAutomation
// GUID:      {F3B79EBD-7D7C-4550-9CBD-BB5179471DFC}
// *********************************************************************//
  IPCsSymbolsEnumerable = interface(IUnknown)
    ['{F3B79EBD-7D7C-4550-9CBD-BB5179471DFC}']
    function GetEnumerator: IPCsSymbolsEnumerator; safecall;
  end;

// *********************************************************************//
// DispIntf:  IPCsSymbolsEnumerableDisp
// Flags:     (320) Dual OleAutomation
// GUID:      {F3B79EBD-7D7C-4550-9CBD-BB5179471DFC}
// *********************************************************************//
  IPCsSymbolsEnumerableDisp = dispinterface
    ['{F3B79EBD-7D7C-4550-9CBD-BB5179471DFC}']
    function GetEnumerator: IPCsSymbolsEnumerator; dispid 101;
  end;

// *********************************************************************//
// Interface: IPCsLine2
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {7653E66C-60EC-467A-90B3-2376CBEFA0FC}
// *********************************************************************//
  IPCsLine2 = interface(IPCsLine)
    ['{7653E66C-60EC-467A-90B3-2376CBEFA0FC}']
    function MoveLinePoint(PointIndex: Integer; Pt: TPCsPoint3D; Options: Integer): WordBool; safecall;
    function MoveLinePointXYZ(PointIndex: Integer; X: Integer; Y: Integer; Z: Integer;
                              Options: Integer): WordBool; safecall;
    function Get_Application: IPCsApplication; safecall;
    function Equals(const Obj: IPCsBase): WordBool; safecall;
    function IsErased: WordBool; safecall;
    property Application: IPCsApplication read Get_Application;
  end;

// *********************************************************************//
// DispIntf:  IPCsLine2Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {7653E66C-60EC-467A-90B3-2376CBEFA0FC}
// *********************************************************************//
  IPCsLine2Disp = dispinterface
    ['{7653E66C-60EC-467A-90B3-2376CBEFA0FC}']
    function MoveLinePoint(PointIndex: Integer; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant;
                           Options: Integer): WordBool; dispid 314;
    function MoveLinePointXYZ(PointIndex: Integer; X: Integer; Y: Integer; Z: Integer;
                              Options: Integer): WordBool; dispid 320;
    property Application: IPCsApplication readonly dispid 321;
    function Equals(const Obj: IPCsBase): WordBool; dispid 322;
    function IsErased: WordBool; dispid 323;
    property Points: PCsPoints readonly dispid 401;
    procedure AddPoint(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 402;
    property IsElectrical: WordBool dispid 301;
    property LineStyle: PsLineStyle dispid 302;
    property Color: PsPenColor dispid 303;
    property PenWidth: Integer dispid 304;
    property Quantity: Double readonly dispid 305;
    property IsFilled: WordBool dispid 306;
    property IsJumperingLink: WordBool dispid 308;
    property IsCableWire: WordBool dispid 307;
    property IsPartLine: WordBool readonly dispid 310;
    property Extended: WordBool dispid 311;
    property IsOpen: WordBool dispid 312;
    property IsSignalsJoined: WordBool readonly dispid 313;
    function AddBendingPoint(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; BendFactor: Double): WordBool; dispid 315;
    procedure ConnectTo(const JoinPoint: IPCsJoinPoint); dispid 316;
    property Pattern: Integer dispid 317;
    property MultilineDistance: Integer dispid 318;
    property Layer: IPCsLayer dispid 319;
    property IsCurvedLine: WordBool dispid 309;
    property Visible: WordBool dispid 201;
    procedure AddPointXYZ(X: Integer; Y: Integer; Z: Integer); dispid 202;
    procedure Delete; dispid 203;
    property LineTexts: IPCsLineTexts readonly dispid 204;
    property Handle: Integer readonly dispid 205;
    property State: Integer dispid 250;
    property ObjectType: PsPCsObjectType readonly dispid 450;
    procedure InitLineTexts; dispid 206;
    function MoveRelative(AOffset: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Figure: IPCsFigure): WordBool; dispid 207;
    function MoveRelativeXYZ(OffsetX: Integer; OffsetY: Integer; OffsetZ: Integer;
                             const Figure: IPCsFigure): WordBool; dispid 208;
    function ChangeDrawingOrder(AAction: psChangeDrawingOrder): WordBool; dispid 209;
  end;

// *********************************************************************//
// DispIntf:  IPCsPageEvents
// Flags:     (0)
// GUID:      {5D80BF9D-FFAC-4B17-B749-2E670C4CAD64}
// *********************************************************************//
  IPCsPageEvents = dispinterface
    ['{5D80BF9D-FFAC-4B17-B749-2E670C4CAD64}']
  end;

// *********************************************************************//
// Interface: IPCsPage2
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {D0465CE5-FBE4-4CDC-9457-991F4C1FBE45}
// *********************************************************************//
  IPCsPage2 = interface(IPCsPage)
    ['{D0465CE5-FBE4-4CDC-9457-991F4C1FBE45}']
    function PlaceReferenceFrame: IPCsReferenceFrame; safecall;
    function SaveToPictureFile(const Filename: WideString; Rect: TPCsRect; Options: Integer): WordBool; safecall;
    function Get_Drawing: IPCsDrawing; safecall;
    function PlaceBusBar(const Filename: WideString; Pt1: TPCsPoint3D; Pt2: TPCsPoint3D): IPCsSymbol; safecall;
    function Equals(const Obj: IPCsBase): WordBool; safecall;
    function Get_Application: IPCsApplication2; safecall;
    function IsErased: WordBool; safecall;
    procedure SetPageHeaderLibtype2(Index: Integer; const LibType: IPCsLibType2); safecall;
    function Get_Pageheaders(Index: Integer): IPCsSymbol; safecall;
    procedure Set_Pageheaders(Index: Integer; const Value: IPCsSymbol); safecall;
    function LoadPageHeader2(Index: Integer; const Filename: WideString): WordBool; safecall;
    function Get_PageHeadersCount: Integer; safecall;
    property Drawing: IPCsDrawing read Get_Drawing;
    property Application: IPCsApplication2 read Get_Application;
    property Pageheaders[Index: Integer]: IPCsSymbol read Get_Pageheaders write Set_Pageheaders;
    property PageHeadersCount: Integer read Get_PageHeadersCount;
  end;

// *********************************************************************//
// DispIntf:  IPCsPage2Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {D0465CE5-FBE4-4CDC-9457-991F4C1FBE45}
// *********************************************************************//
  IPCsPage2Disp = dispinterface
    ['{D0465CE5-FBE4-4CDC-9457-991F4C1FBE45}']
    function PlaceReferenceFrame: IPCsReferenceFrame; dispid 417;
    function SaveToPictureFile(const Filename: WideString; Rect: {NOT_OLEAUTO(TPCsRect)}OleVariant;
                               Options: Integer): WordBool; dispid 418;
    property Drawing: IPCsDrawing readonly dispid 411;
    function PlaceBusBar(const Filename: WideString; Pt1: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant;
                         Pt2: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant): IPCsSymbol; dispid 419;
    function Equals(const Obj: IPCsBase): WordBool; dispid 420;
    property Application: IPCsApplication2 readonly dispid 421;
    function IsErased: WordBool; dispid 422;
    procedure SetPageHeaderLibtype2(Index: Integer; const LibType: IPCsLibType2); dispid 423;
    property Pageheaders[Index: Integer]: IPCsSymbol dispid 424;
    function LoadPageHeader2(Index: Integer; const Filename: WideString): WordBool; dispid 425;
    property PageHeadersCount: Integer readonly dispid 426;
    property Symbols: IPCsSymbols readonly dispid 401;
    property PageNumber: WideString dispid 402;
    property PageSetup: IPCsPageSetup readonly dispid 403;
    property PageHeader: IPCsSymbol dispid 404;
    function LoadPageHeader(const Filename: WideString): WordBool; dispid 405;
    property IsEmpty: WordBool readonly dispid 406;
    property PageType: psPageType dispid 407;
    property PageFunction: psPageFunction dispid 408;
    property Origo: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant dispid 409;
    property Index: Integer dispid 410;
    procedure CreateJoins; dispid 412;
    procedure JoinParkedLines; dispid 413;
    procedure JoinConnectionsToCrossingLines; dispid 414;
    function PlaceSymbol(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Filename: WideString): IPCsSymbol; dispid 415;
    property PageReference: IPCsPageReference readonly dispid 416;
    property Scale: Single dispid 301;
    function GetMasterPage: IPCsPage; dispid 311;
    property ReadingDirection: Integer dispid 304;
    procedure Draw(DeviceContext: LongWord; Rect: {NOT_OLEAUTO(TPCsRect)}OleVariant); dispid 305;
    procedure DrawOntoWindow(Wnd: TWindowHandle; Rect: {NOT_OLEAUTO(TPCsRect)}OleVariant); dispid 306;
    property ExternalObjects: IPCsExternalObjects readonly dispid 307;
    function GetRect(AFlags: psPageRectFlag): {NOT_OLEAUTO(TPCsRect)}OleVariant; dispid 302;
    property PageData[Index: Integer]: WideString dispid 308;
    property PageDataCount: Integer readonly dispid 309;
    procedure Delete; dispid 310;
    procedure PrintOut; dispid 312;
    procedure Move(Index: Integer); dispid 313;
    function GetVirtualsAsSymbols: IPCsSymbols; dispid 314;
    function PlaceSubDrawing(const SubDrawing: IPCsSubDrawing; Group: Integer; X: Integer;
                             Y: Integer; Z: Integer; const Layer: IPCsLayer;
                             Options: psPlaceSubDrawingOption): WordBool; dispid 315;
    procedure SetPageHeaderLibtype(const LibType: IPCsLibType); dispid 316;
    property Joins: IPCsJoins readonly dispid 500;
    procedure OffsetMoveObjects(AX: Integer; AY: Integer; AZ: Integer); dispid 317;
    property RefDesLocation: WideString dispid 318;
    property RefDesFunction: WideString dispid 319;
    function UpdateFromDatabase(Reserved: Integer): WordBool; dispid 320;
    property Lines: IPCsLines readonly dispid 201;
    property Connections: IPCsLibTypeConnections readonly dispid 202;
    property Arcs: IPCsArcs readonly dispid 203;
    property Texts: IPCsTexts readonly dispid 204;
    property Datafields: IPCsDatafields readonly dispid 205;
    property Title: WideString dispid 303;
    property TitleLC[LCID: Integer]: WideString dispid 206;
    property States: OleVariant readonly dispid 207;
    property SymbolTexts: IPCsSymbolTexts readonly dispid 208;
  end;

// *********************************************************************//
// Interface: IPCsArc2
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {02F30A6A-EC65-4B42-9189-31946D8D6CB0}
// *********************************************************************//
  IPCsArc2 = interface(IPCsArc)
    ['{02F30A6A-EC65-4B42-9189-31946D8D6CB0}']
    function Get_RGBColor: TRGBColor; safecall;
    procedure Set_RGBColor(Value: TRGBColor); safecall;
    property RGBColor: TRGBColor read Get_RGBColor write Set_RGBColor;
  end;

// *********************************************************************//
// DispIntf:  IPCsArc2Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {02F30A6A-EC65-4B42-9189-31946D8D6CB0}
// *********************************************************************//
  IPCsArc2Disp = dispinterface
    ['{02F30A6A-EC65-4B42-9189-31946D8D6CB0}']
    property RGBColor: TRGBColor dispid 402;
    property IsFilled: WordBool dispid 401;
    property Radius: Integer dispid 403;
    property StartAngle: Integer dispid 404;
    property EndAngle: Integer dispid 405;
    property PenWidth: Integer dispid 407;
    property _EllipseFactor: Integer dispid 408;
    property Color: PsPenColor dispid 409;
    property Layer: IPCsLayer dispid 410;
    property IsCircle: WordBool readonly dispid 411;
    property Rotation: Integer dispid 412;
    property EllipseFactor: Double dispid 413;
    procedure Delete; dispid 414;
    property Visible: WordBool dispid 201;
    property Position: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant dispid 202;
    property LStatus: Integer dispid 203;
    function Move(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Figure: IPCsFigure): WordBool; dispid 204;
    procedure Rotate(DeltaAngle: Integer; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 205;
    procedure Scale(Factor: Single; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 208;
    property Handle: LongWord readonly dispid 206;
    procedure GetPosition(out X: Integer; out Y: Integer; out Z: Integer); dispid 207;
    procedure SetPosition(X: Integer; Y: Integer; Z: Integer); dispid 209;
    property State: Integer dispid 250;
    function ChangeDrawingOrder(AAction: psChangeDrawingOrder): WordBool; dispid 210;
  end;

// *********************************************************************//
// Interface: IPCsStatusObject
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {E78749F9-B862-46B4-B3D9-FDC143E67FDF}
// *********************************************************************//
  IPCsStatusObject = interface(IDispatch)
    ['{E78749F9-B862-46B4-B3D9-FDC143E67FDF}']
    function Get_InternalOleID: Integer; safecall;
    procedure Set_InternalOleID(Value: Integer); safecall;
    function Select(Options: psSelectOptions): WordBool; safecall;
    function Get_Selected: WordBool; safecall;
    procedure Set_Selected(Value: WordBool); safecall;
    function Get_Highlightning: psHighlighting; safecall;
    procedure Set_Highlightning(Value: psHighlighting); safecall;
    property InternalOleID: Integer read Get_InternalOleID write Set_InternalOleID;
    property Selected: WordBool read Get_Selected write Set_Selected;
    property Highlightning: psHighlighting read Get_Highlightning write Set_Highlightning;
  end;

// *********************************************************************//
// DispIntf:  IPCsStatusObjectDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {E78749F9-B862-46B4-B3D9-FDC143E67FDF}
// *********************************************************************//
  IPCsStatusObjectDisp = dispinterface
    ['{E78749F9-B862-46B4-B3D9-FDC143E67FDF}']
    property InternalOleID: Integer dispid 201;
    function Select(Options: psSelectOptions): WordBool; dispid 300;
    property Selected: WordBool dispid 302;
    property Highlightning: psHighlighting dispid 202;
  end;

// *********************************************************************//
// Interface: IPCsConnection2
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {94CF8C33-E3F3-49CE-B9D9-D695918670EF}
// *********************************************************************//
  IPCsConnection2 = interface(IPCsConnection)
    ['{94CF8C33-E3F3-49CE-B9D9-D695918670EF}']
    function IsSamePin(const ASymbol: IPCsSymbol; AConnectionIndex: Integer;
                       Options: psIsSameComponentOptions): WordBool; safecall;
    function IsConnectionPair(const OtherConnection: IPCsConnection): WordBool; safecall;
    function IsMetapin(Options: psMetaconnectivityOptions): WordBool; safecall;
    function Get_Symbol: IPCsSymbol; safecall;
    property Symbol: IPCsSymbol read Get_Symbol;
  end;

// *********************************************************************//
// DispIntf:  IPCsConnection2Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {94CF8C33-E3F3-49CE-B9D9-D695918670EF}
// *********************************************************************//
  IPCsConnection2Disp = dispinterface
    ['{94CF8C33-E3F3-49CE-B9D9-D695918670EF}']
    function IsSamePin(const ASymbol: IPCsSymbol; AConnectionIndex: Integer;
                       Options: psIsSameComponentOptions): WordBool; dispid 501;
    function IsConnectionPair(const OtherConnection: IPCsConnection): WordBool; dispid 502;
    function IsMetapin(Options: psMetaconnectivityOptions): WordBool; dispid 503;
    property Symbol: IPCsSymbol readonly dispid 504;
    property IsElectrical: WordBool dispid 401;
    property IOStatus: {NOT_OLEAUTO(TPCsIOStatus)}OleVariant dispid 402;
    property IsJumperingLink: WordBool dispid 403;
    property IsCableScreen: WordBool dispid 404;
    property ConnectionTexts: IPCsConnectionTexts readonly dispid 405;
    property Layer: IPCsLayer dispid 406;
    property PinName: WideString dispid 407;
    procedure InitConnectionTexts(const Drawing: IPCsDrawing); dispid 408;
    procedure Delete; dispid 409;
    property WithReference: WordBool dispid 410;
    property Reference: IPCsReference readonly dispid 411;
    property Visible: WordBool dispid 201;
    property Position: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant dispid 202;
    property LStatus: Integer dispid 203;
    function Move(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Figure: IPCsFigure): WordBool; dispid 204;
    procedure Rotate(DeltaAngle: Integer; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 205;
    procedure Scale(Factor: Single; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 208;
    property Handle: LongWord readonly dispid 206;
    procedure GetPosition(out X: Integer; out Y: Integer; out Z: Integer); dispid 207;
    procedure SetPosition(X: Integer; Y: Integer; Z: Integer); dispid 209;
    property State: Integer dispid 250;
    function ChangeDrawingOrder(AAction: psChangeDrawingOrder): WordBool; dispid 210;
  end;

// *********************************************************************//
// Interface: IPCsComponentDatabase2
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {FE0869BA-D134-4ED5-80ED-E57EFA8A4CF6}
// *********************************************************************//
  IPCsComponentDatabase2 = interface(IPCsComponentDatabase)
    ['{FE0869BA-D134-4ED5-80ED-E57EFA8A4CF6}']
    function Get_FieldNames(AIndex: Integer): WideString; safecall;
    function Get_FieldCount: Integer; safecall;
    function LookupByFieldnames(const AKeyField: WideString; AKeyValue: OleVariant;
                                const AResultFields: WideString): OleVariant; safecall;
    function Lookup(AKeyField: psSetupDatabaseFieldTypes; AKeyValue: OleVariant;
                    AResultFields: OleVariant): OleVariant; safecall;
    function FieldExists(const AFieldName: WideString): WordBool; safecall;
    property FieldNames[AIndex: Integer]: WideString read Get_FieldNames;
    property FieldCount: Integer read Get_FieldCount;
  end;

// *********************************************************************//
// DispIntf:  IPCsComponentDatabase2Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {FE0869BA-D134-4ED5-80ED-E57EFA8A4CF6}
// *********************************************************************//
  IPCsComponentDatabase2Disp = dispinterface
    ['{FE0869BA-D134-4ED5-80ED-E57EFA8A4CF6}']
    property FieldNames[AIndex: Integer]: WideString readonly dispid 207;
    property FieldCount: Integer readonly dispid 208;
    function LookupByFieldnames(const AKeyField: WideString; AKeyValue: OleVariant;
                                const AResultFields: WideString): OleVariant; dispid 209;
    function Lookup(AKeyField: psSetupDatabaseFieldTypes; AKeyValue: OleVariant;
                    AResultFields: OleVariant): OleVariant; dispid 210;
    function FieldExists(const AFieldName: WideString): WordBool; dispid 301;
    property Name: Integer readonly dispid 201;
    property Active: WordBool readonly dispid 202;
    property TableComponents: WideString readonly dispid 203;
    property DatabaseAlias: WideString readonly dispid 204;
    property DatabaseFile: WideString readonly dispid 205;
    property DatabaseString: WideString readonly dispid 206;
  end;

// *********************************************************************//
// Interface: IPCsGlobal
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {C6BF0C78-7D73-4D6D-B833-743580EE5B5D}
// *********************************************************************//
  IPCsGlobal = interface(IDispatch)
    ['{C6BF0C78-7D73-4D6D-B833-743580EE5B5D}']
  end;

// *********************************************************************//
// DispIntf:  IPCsGlobalDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {C6BF0C78-7D73-4D6D-B833-743580EE5B5D}
// *********************************************************************//
  IPCsGlobalDisp = dispinterface
    ['{C6BF0C78-7D73-4D6D-B833-743580EE5B5D}']
  end;

// *********************************************************************//
// Interface: IPCsGlobal2
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {200950A3-3BE4-415A-B5C9-10D3D5A340C8}
// *********************************************************************//
  IPCsGlobal2 = interface(IPCsGlobal)
    ['{200950A3-3BE4-415A-B5C9-10D3D5A340C8}']
    function BuildReferenceDesignationString(const FunctionalAspect: WideString;
                                             const LocationAspect: WideString;
                                             const ProductAspect: WideString;
                                             const ComponentName: WideString;
                                             const FormatString: WideString): WideString; safecall;
  end;

// *********************************************************************//
// DispIntf:  IPCsGlobal2Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {200950A3-3BE4-415A-B5C9-10D3D5A340C8}
// *********************************************************************//
  IPCsGlobal2Disp = dispinterface
    ['{200950A3-3BE4-415A-B5C9-10D3D5A340C8}']
    function BuildReferenceDesignationString(const FunctionalAspect: WideString;
                                             const LocationAspect: WideString;
                                             const ProductAspect: WideString;
                                             const ComponentName: WideString;
                                             const FormatString: WideString): WideString; dispid 301;
  end;

// *********************************************************************//
// Interface: IPCsExternalObjects2
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {E42E13EE-7A40-4C50-9FDE-F41360A4A199}
// *********************************************************************//
  IPCsExternalObjects2 = interface(IPCsExternalObjects)
    ['{E42E13EE-7A40-4C50-9FDE-F41360A4A199}']
    function Get_Count: Integer; safecall;
    function Get_Item(Index: Integer): IPCsExternalObject2; safecall;
    procedure LoadPicture(Position: TPCsPoint3D; const Filename: WideString; Size: TPCsPoint2D;
                          Embedded: WordBool); safecall;
    property Count: Integer read Get_Count;
    property Item[Index: Integer]: IPCsExternalObject2 read Get_Item;
  end;

// *********************************************************************//
// DispIntf:  IPCsExternalObjects2Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {E42E13EE-7A40-4C50-9FDE-F41360A4A199}
// *********************************************************************//
  IPCsExternalObjects2Disp = dispinterface
    ['{E42E13EE-7A40-4C50-9FDE-F41360A4A199}']
    property Count: Integer readonly dispid 301;
    property Item[Index: Integer]: IPCsExternalObject2 readonly dispid 302;
    procedure LoadPicture(Position: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant;
                          const Filename: WideString; Size: {NOT_OLEAUTO(TPCsPoint2D)}OleVariant;
                          Embedded: WordBool); dispid 303;
    procedure AddMetafile(Position: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant;
                          EnhMetafileHandle: Integer; Size: {NOT_OLEAUTO(TPCsPoint2D)}OleVariant); dispid 202;
  end;

// *********************************************************************//
// Interface: IPCsDatafields2
// Flags:     (320) Dual OleAutomation
// GUID:      {4268FFB3-BE93-47E6-A20D-E898BFEE0467}
// *********************************************************************//
  IPCsDatafields2 = interface(IPCsDatafields)
    ['{4268FFB3-BE93-47E6-A20D-E898BFEE0467}']
    function Get_ValueByIndex(Index: Integer): WideString; safecall;
    procedure Set_ValueByIndex(Index: Integer; const Value: WideString); safecall;
    function Get_ValueByName(const Name: WideString): WideString; safecall;
    procedure Set_ValueByName(const Name: WideString; const Value: WideString); safecall;
    function Get_NameByIndex(Index: Integer): WideString; safecall;
    procedure Set_NameByIndex(Index: Integer; const Name: WideString); safecall;
    property ValueByIndex[Index: Integer]: WideString read Get_ValueByIndex write Set_ValueByIndex;
    property ValueByName[const Name: WideString]: WideString read Get_ValueByName write Set_ValueByName;
    property NameByIndex[Index: Integer]: WideString read Get_NameByIndex write Set_NameByIndex;
  end;

// *********************************************************************//
// DispIntf:  IPCsDatafields2Disp
// Flags:     (320) Dual OleAutomation
// GUID:      {4268FFB3-BE93-47E6-A20D-E898BFEE0467}
// *********************************************************************//
  IPCsDatafields2Disp = dispinterface
    ['{4268FFB3-BE93-47E6-A20D-E898BFEE0467}']
    property ValueByIndex[Index: Integer]: WideString dispid 201;
    property ValueByName[const Name: WideString]: WideString dispid 202;
    property NameByIndex[Index: Integer]: WideString dispid 203;
    function Add(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Value: WideString): IPCsDatafield; dispid 101;
    property Item[Index: Integer]: IPCsDatafield readonly dispid 0; default;
    property Count: Integer readonly dispid 102;
    function FindByValue(const Value: WideString): IPCsDatafield; dispid 103;
    property _NewEnum: IEnumVARIANT readonly dispid -4;
  end;

// *********************************************************************//
// Interface: IPCsNameText2
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {32E89920-84E7-474A-BC7E-E53AE9AAC79F}
// *********************************************************************//
  IPCsNameText2 = interface(IPCsNameText)
    ['{32E89920-84E7-474A-BC7E-E53AE9AAC79F}']
    function Get_DisplayText2: WideString; safecall;
    function Equals(const Obj: IPCsBase): WordBool; safecall;
    function Get_Application: IPCsApplication2; safecall;
    function IsErased: WordBool; safecall;
    property DisplayText2: WideString read Get_DisplayText2;
    property Application: IPCsApplication2 read Get_Application;
  end;

// *********************************************************************//
// DispIntf:  IPCsNameText2Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {32E89920-84E7-474A-BC7E-E53AE9AAC79F}
// *********************************************************************//
  IPCsNameText2Disp = dispinterface
    ['{32E89920-84E7-474A-BC7E-E53AE9AAC79F}']
    property DisplayText2: WideString readonly dispid 501;
    function Equals(const Obj: IPCsBase): WordBool; dispid 502;
    property Application: IPCsApplication2 readonly dispid 503;
    function IsErased: WordBool; dispid 504;
    property RefDesFunction: WideString dispid 409;
    property RefDesLocation: WideString dispid 410;
    property RefDesOptions: psSymbolRefDesOption dispid 411;
    property Value: WideString dispid 401;
    property Origin: PsTextOrigin dispid 402;
    property Font: IPCsTextFont readonly dispid 403;
    property TextType: PsTextOrigin readonly dispid 404;
    property FieldWidth: Integer dispid 405;
    property WrapText: WordBool dispid 406;
    property LinkID: Integer dispid 407;
    property PositionExt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant dispid 408;
    property Color: PsPenColor dispid 301;
    property Layer: IPCsLayer dispid 302;
    property DisplayText: WideString readonly dispid 303;
    property LineCount: Integer readonly dispid 304;
    property Rotation: Integer dispid 305;
    property IsBlank: WordBool readonly dispid 306;
    function GetRect(AFlags: psTextRectFlags): {NOT_OLEAUTO(TPCsRect)}OleVariant; dispid 307;
    procedure Delete; dispid 308;
    property ObjectType: PsPCsObjectType readonly dispid 450;
    property NotTranslated: WordBool dispid 309;
    property Visible: WordBool dispid 201;
    property Position: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant dispid 202;
    property LStatus: Integer dispid 203;
    function Move(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Figure: IPCsFigure): WordBool; dispid 204;
    procedure Rotate(DeltaAngle: Integer; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 205;
    procedure Scale(Factor: Single; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 208;
    property Handle: LongWord readonly dispid 206;
    procedure GetPosition(out X: Integer; out Y: Integer; out Z: Integer); dispid 207;
    procedure SetPosition(X: Integer; Y: Integer; Z: Integer); dispid 209;
    property State: Integer dispid 250;
    function ChangeDrawingOrder(AAction: psChangeDrawingOrder): WordBool; dispid 210;
  end;

// *********************************************************************//
// Interface: IPCsReferenceDesignations
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {42F4159D-C012-4641-9D20-A740E69BF862}
// *********************************************************************//
  IPCsReferenceDesignations = interface(IDispatch)
    ['{42F4159D-C012-4641-9D20-A740E69BF862}']
    procedure Add(Aspect: psReferenceDesignationAspect; const Designation: WideString;
                  const Description: WideString); safecall;
    function Get_Location(Index: Integer): IPCsReferenceDesignation; safecall;
    function Get_Func(Index: Integer): IPCsReferenceDesignation; safecall;
    function Get_LocationCount: Integer; safecall;
    function Get_FunctionCount: Integer; safecall;
    function Find(Aspect: psReferenceDesignationAspect; const Designation: WideString): IPCsReferenceDesignation; safecall;
    property Location[Index: Integer]: IPCsReferenceDesignation read Get_Location;
    property Func[Index: Integer]: IPCsReferenceDesignation read Get_Func;
    property LocationCount: Integer read Get_LocationCount;
    property FunctionCount: Integer read Get_FunctionCount;
  end;

// *********************************************************************//
// DispIntf:  IPCsReferenceDesignationsDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {42F4159D-C012-4641-9D20-A740E69BF862}
// *********************************************************************//
  IPCsReferenceDesignationsDisp = dispinterface
    ['{42F4159D-C012-4641-9D20-A740E69BF862}']
    procedure Add(Aspect: psReferenceDesignationAspect; const Designation: WideString;
                  const Description: WideString); dispid 201;
    property Location[Index: Integer]: IPCsReferenceDesignation readonly dispid 202;
    property Func[Index: Integer]: IPCsReferenceDesignation readonly dispid 203;
    property LocationCount: Integer readonly dispid 204;
    property FunctionCount: Integer readonly dispid 205;
    function Find(Aspect: psReferenceDesignationAspect; const Designation: WideString): IPCsReferenceDesignation; dispid 206;
  end;

// *********************************************************************//
// Interface: IPCsReferenceDesignations2
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {3FB35D3D-29E2-48D9-B6DA-687710CC743C}
// *********************************************************************//
  IPCsReferenceDesignations2 = interface(IPCsReferenceDesignations)
    ['{3FB35D3D-29E2-48D9-B6DA-687710CC743C}']
    function Get_Aspects(Aspect: psReferenceDesignationAspect): IPCsReferenceDesignationList; safecall;
    property Aspects[Aspect: psReferenceDesignationAspect]: IPCsReferenceDesignationList read Get_Aspects;
  end;

// *********************************************************************//
// DispIntf:  IPCsReferenceDesignations2Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {3FB35D3D-29E2-48D9-B6DA-687710CC743C}
// *********************************************************************//
  IPCsReferenceDesignations2Disp = dispinterface
    ['{3FB35D3D-29E2-48D9-B6DA-687710CC743C}']
    property Aspects[Aspect: psReferenceDesignationAspect]: IPCsReferenceDesignationList readonly dispid 401;
    procedure Add(Aspect: psReferenceDesignationAspect; const Designation: WideString;
                  const Description: WideString); dispid 201;
    property Location[Index: Integer]: IPCsReferenceDesignation readonly dispid 202;
    property Func[Index: Integer]: IPCsReferenceDesignation readonly dispid 203;
    property LocationCount: Integer readonly dispid 204;
    property FunctionCount: Integer readonly dispid 205;
    function Find(Aspect: psReferenceDesignationAspect; const Designation: WideString): IPCsReferenceDesignation; dispid 206;
  end;

// *********************************************************************//
// Interface: IPCsReferenceDesignation
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {03DA2962-6DA8-47A0-85BA-279314C3E075}
// *********************************************************************//
  IPCsReferenceDesignation = interface(IDispatch)
    ['{03DA2962-6DA8-47A0-85BA-279314C3E075}']
    function Get_Designation: WideString; safecall;
    procedure Set_Designation(const Value: WideString); safecall;
    function Get_Description: WideString; safecall;
    procedure Set_Description(const Value: WideString); safecall;
    property Designation: WideString read Get_Designation write Set_Designation;
    property Description: WideString read Get_Description write Set_Description;
  end;

// *********************************************************************//
// DispIntf:  IPCsReferenceDesignationDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {03DA2962-6DA8-47A0-85BA-279314C3E075}
// *********************************************************************//
  IPCsReferenceDesignationDisp = dispinterface
    ['{03DA2962-6DA8-47A0-85BA-279314C3E075}']
    property Designation: WideString dispid 201;
    property Description: WideString dispid 202;
  end;

// *********************************************************************//
// Interface: IPCsSymbolFolderAliases
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {F70942D5-008E-4E4D-B415-E305833B7EF4}
// *********************************************************************//
  IPCsSymbolFolderAliases = interface(IDispatch)
    ['{F70942D5-008E-4E4D-B415-E305833B7EF4}']
    function Get_Aliases(Index: Integer): WideString; safecall;
    procedure Set_Aliases(Index: Integer; const Value: WideString); safecall;
    function Get_Folders(Index: Integer): WideString; safecall;
    procedure Set_Folders(Index: Integer; const Value: WideString); safecall;
    function Get_Count: Integer; safecall;
    procedure Add(const Alias: WideString; const Path: WideString); safecall;
    function GetPathOfFile(const Filename: WideString; const Drawing: IPCsDrawing): WideString; safecall;
    property Aliases[Index: Integer]: WideString read Get_Aliases write Set_Aliases;
    property Folders[Index: Integer]: WideString read Get_Folders write Set_Folders;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  IPCsSymbolFolderAliasesDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {F70942D5-008E-4E4D-B415-E305833B7EF4}
// *********************************************************************//
  IPCsSymbolFolderAliasesDisp = dispinterface
    ['{F70942D5-008E-4E4D-B415-E305833B7EF4}']
    property Aliases[Index: Integer]: WideString dispid 201;
    property Folders[Index: Integer]: WideString dispid 202;
    property Count: Integer readonly dispid 203;
    procedure Add(const Alias: WideString; const Path: WideString); dispid 204;
    function GetPathOfFile(const Filename: WideString; const Drawing: IPCsDrawing): WideString; dispid 205;
  end;

// *********************************************************************//
// Interface: IPCsApplication3
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {DEE2514B-736F-4A32-AD5C-35E6C7373407}
// *********************************************************************//
  IPCsApplication3 = interface(IPCsApplication2)
    ['{DEE2514B-736F-4A32-AD5C-35E6C7373407}']
    function Get_Codepage: Integer; safecall;
    procedure LockUserinterface(WindowHandle: TWindowHandle; Lock: WordBool); safecall;
    procedure SetWaitCursor(out Restore: IUnknown); safecall;
    property Codepage: Integer read Get_Codepage;
  end;

// *********************************************************************//
// DispIntf:  IPCsApplication3Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {DEE2514B-736F-4A32-AD5C-35E6C7373407}
// *********************************************************************//
  IPCsApplication3Disp = dispinterface
    ['{DEE2514B-736F-4A32-AD5C-35E6C7373407}']
    property Codepage: Integer readonly dispid 401;
    procedure LockUserinterface(WindowHandle: TWindowHandle; Lock: WordBool); dispid 407;
    procedure SetWaitCursor(out Restore: IUnknown); dispid 415;
    function GetDocumentFromHandle(Handle: Integer): IPCsDocument; dispid 301;
    function CreateObject(ClassID: {NOT_OLEAUTO(TGUID)}OleVariant; const Owner: IUnknown): IUnknown; dispid 302;
    function Equals(const Obj: IPCsBase): WordBool; dispid 303;
    property Application: IPCsApplication2 readonly dispid 304;
    property Documents: IPCsDocuments readonly dispid 400;
    property AvailableFonts: OleVariant readonly dispid 402;
    property PenColors[PenColor: PsPenColor]: OLE_COLOR readonly dispid 403;
    property Preferences: IPCsApplicationPreferences readonly dispid 404;
    property SymbolLibraryMenu: IPCsSymbolMenu readonly dispid 405;
    property StatusBar: WideString dispid 406;
    property Visible: WordBool dispid 408;
    procedure Quit; dispid 409;
    procedure DontShutDownOnLastRelease; dispid 410;
    function CreateVirtualPage: IPCsPage; dispid 411;
    property ActiveComponentDatabase: IPCsComponentDatabase readonly dispid 412;
    function ExecuteModule(const Filename: WideString; const ParameterString: WideString): WordBool; dispid 413;
    procedure Redraw; dispid 414;
    property ActiveDocument: IPCsDocument dispid 201;
    procedure ZoomAll; dispid 202;
    function LoadString(Identifier: Integer): WideString; dispid 203;
    function SetOleStatusbarText(PanelIndex: Integer; const Text: WideString; Width: Integer;
                                 Alignment: Shortint): TWindowHandle; dispid 204;
    property Dialogs[Kind: psDialogKind]: IPCsDialog readonly dispid 205;
    procedure Activate; dispid 206;
    property ExplorerWindow: IPCsExplorerWindow readonly dispid 207;
    function CreateNewLibType: IPCsLibType; dispid 208;
  end;

// *********************************************************************//
// Interface: IPCsDialog
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {C7485B3E-E4D1-49C0-AD43-BD4E390FC792}
// *********************************************************************//
  IPCsDialog = interface(IDispatch)
    ['{C7485B3E-E4D1-49C0-AD43-BD4E390FC792}']
    function Execute: Integer; safecall;
  end;

// *********************************************************************//
// DispIntf:  IPCsDialogDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {C7485B3E-E4D1-49C0-AD43-BD4E390FC792}
// *********************************************************************//
  IPCsDialogDisp = dispinterface
    ['{C7485B3E-E4D1-49C0-AD43-BD4E390FC792}']
    function Execute: Integer; dispid 201;
  end;

// *********************************************************************//
// DispIntf:  IPCsDialogEvents
// Flags:     (0)
// GUID:      {1370ECF3-341D-43F9-9B70-0C111D611035}
// *********************************************************************//
  IPCsDialogEvents = dispinterface
    ['{1370ECF3-341D-43F9-9B70-0C111D611035}']
  end;

// *********************************************************************//
// Interface: IPCsFileDialog
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {5C3A5FE5-91C7-44ED-BE25-EEE64F549345}
// *********************************************************************//
  IPCsFileDialog = interface(IPCsDialog)
    ['{5C3A5FE5-91C7-44ED-BE25-EEE64F549345}']
    function Get_Filename: WideString; safecall;
    procedure Set_Filename(const Value: WideString); safecall;
    function Get_Files: OleVariant; safecall;
    procedure Set_Files(Value: OleVariant); safecall;
    function Get_InitialDir: WideString; safecall;
    procedure Set_InitialDir(const Value: WideString); safecall;
    function Get_Title: WideString; safecall;
    procedure Set_Title(const Value: WideString); safecall;
    property Filename: WideString read Get_Filename write Set_Filename;
    property Files: OleVariant read Get_Files write Set_Files;
    property InitialDir: WideString read Get_InitialDir write Set_InitialDir;
    property Title: WideString read Get_Title write Set_Title;
  end;

// *********************************************************************//
// DispIntf:  IPCsFileDialogDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {5C3A5FE5-91C7-44ED-BE25-EEE64F549345}
// *********************************************************************//
  IPCsFileDialogDisp = dispinterface
    ['{5C3A5FE5-91C7-44ED-BE25-EEE64F549345}']
    property Filename: WideString dispid 301;
    property Files: OleVariant dispid 302;
    property InitialDir: WideString dispid 303;
    property Title: WideString dispid 304;
    function Execute: Integer; dispid 201;
  end;

// *********************************************************************//
// DispIntf:  IPCsFileDialogEvents
// Flags:     (0)
// GUID:      {6525E6D9-BCD5-4E09-AE66-E5933DA8227A}
// *********************************************************************//
  IPCsFileDialogEvents = dispinterface
    ['{6525E6D9-BCD5-4E09-AE66-E5933DA8227A}']
  end;

// *********************************************************************//
// Interface: IPCsLineTexts
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0423A4FD-AD69-406C-B337-2999E0FA7AE0}
// *********************************************************************//
  IPCsLineTexts = interface(IPCsSymbolTexts)
    ['{0423A4FD-AD69-406C-B337-2999E0FA7AE0}']
  end;

// *********************************************************************//
// DispIntf:  IPCsLineTextsDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0423A4FD-AD69-406C-B337-2999E0FA7AE0}
// *********************************************************************//
  IPCsLineTextsDisp = dispinterface
    ['{0423A4FD-AD69-406C-B337-2999E0FA7AE0}']
    property NameText: IPCsNameText readonly dispid 301;
    property ArticleText: IPCsText readonly dispid 302;
    property TypeText: IPCsText readonly dispid 303;
    property FunctionText: IPCsText readonly dispid 304;
    function FindByType(AType: PsTextType): IPCsText; dispid 305;
    property _NewEnum: IEnumVARIANT readonly dispid -4;
    property Item[Index: Integer]: IPCsText readonly dispid 0; default;
    property Count: Integer readonly dispid 203;
    function Add(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Value: WideString): IPCsText; dispid 204;
    function FindByValue(const Value: WideString): IPCsText; dispid 205;
    procedure Remove(Index: Integer); dispid 201;
    function AddXYZ(X: Integer; Y: Integer; Z: Integer; const Value: WideString): IPCsText; dispid 202;
  end;

// *********************************************************************//
// Interface: IPCsLists
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {A7AEB7AD-3A64-43FF-9836-C9663D974EF2}
// *********************************************************************//
  IPCsLists = interface(IDispatch)
    ['{A7AEB7AD-3A64-43FF-9836-C9663D974EF2}']
    function CreateCableList(const ParamStr: WideString): IPCsList; safecall;
    function CreateTerminalList(const ParamStr: WideString): IPCsList; safecall;
    function CreatePartList(const ParamStr: WideString): IPCsList; safecall;
    function CreateComponentList(const ParamStr: WideString): IPCsList; safecall;
    procedure ListToFile(const NewListFile: WideString; ListType: psListType;
                         const FormatFile: WideString; const Options: WideString); safecall;
  end;

// *********************************************************************//
// DispIntf:  IPCsListsDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {A7AEB7AD-3A64-43FF-9836-C9663D974EF2}
// *********************************************************************//
  IPCsListsDisp = dispinterface
    ['{A7AEB7AD-3A64-43FF-9836-C9663D974EF2}']
    function CreateCableList(const ParamStr: WideString): IPCsList; dispid 201;
    function CreateTerminalList(const ParamStr: WideString): IPCsList; dispid 202;
    function CreatePartList(const ParamStr: WideString): IPCsList; dispid 203;
    function CreateComponentList(const ParamStr: WideString): IPCsList; dispid 204;
    procedure ListToFile(const NewListFile: WideString; ListType: psListType;
                         const FormatFile: WideString; const Options: WideString); dispid 205;
  end;

// *********************************************************************//
// Interface: IPCsLists2
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {5321BF67-C846-444E-ABB6-36CF6CA8C0F3}
// *********************************************************************//
  IPCsLists2 = interface(IDispatch)
    ['{5321BF67-C846-444E-ABB6-36CF6CA8C0F3}']
    function CreateConnectionList(const AParamStr: WideString): IPCsList2; safecall;
    procedure UpdateLists(PageFunction: psPageFunction); safecall;
  end;

// *********************************************************************//
// DispIntf:  IPCsLists2Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {5321BF67-C846-444E-ABB6-36CF6CA8C0F3}
// *********************************************************************//
  IPCsLists2Disp = dispinterface
    ['{5321BF67-C846-444E-ABB6-36CF6CA8C0F3}']
    function CreateConnectionList(const AParamStr: WideString): IPCsList2; dispid 301;
    procedure UpdateLists(PageFunction: psPageFunction); dispid 302;
  end;

// *********************************************************************//
// Interface: IPCsList
// Flags:     (4432) Hidden Dual OleAutomation Dispatchable
// GUID:      {32055A0A-D65C-4A85-A964-E00403423559}
// *********************************************************************//
  IPCsList = interface(IDispatch)
    ['{32055A0A-D65C-4A85-A964-E00403423559}']
    function Get_ListType: psListType; safecall;
    function Get_Count: Integer; safecall;
    function Get_Item(Index: Integer): IPCsListConnData; safecall;
    property ListType: psListType read Get_ListType;
    property Count: Integer read Get_Count;
    property Item[Index: Integer]: IPCsListConnData read Get_Item; default;
  end;

// *********************************************************************//
// DispIntf:  IPCsListDisp
// Flags:     (4432) Hidden Dual OleAutomation Dispatchable
// GUID:      {32055A0A-D65C-4A85-A964-E00403423559}
// *********************************************************************//
  IPCsListDisp = dispinterface
    ['{32055A0A-D65C-4A85-A964-E00403423559}']
    property ListType: psListType readonly dispid 201;
    property Count: Integer readonly dispid 202;
    property Item[Index: Integer]: IPCsListConnData readonly dispid 0; default;
  end;

// *********************************************************************//
// Interface: IPCsList2
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {F38CB7E9-64C0-47D6-A484-40E4378F8623}
// *********************************************************************//
  IPCsList2 = interface(IPCsList)
    ['{F38CB7E9-64C0-47D6-A484-40E4378F8623}']
    procedure SaveToFile(const Filename: WideString; Options: Integer); safecall;
    function Equals(const Obj: IPCsBase): WordBool; safecall;
    function Get_Application: IPCsApplication2; safecall;
    property Application: IPCsApplication2 read Get_Application;
  end;

// *********************************************************************//
// DispIntf:  IPCsList2Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {F38CB7E9-64C0-47D6-A484-40E4378F8623}
// *********************************************************************//
  IPCsList2Disp = dispinterface
    ['{F38CB7E9-64C0-47D6-A484-40E4378F8623}']
    procedure SaveToFile(const Filename: WideString; Options: Integer); dispid 301;
    function Equals(const Obj: IPCsBase): WordBool; dispid 302;
    property Application: IPCsApplication2 readonly dispid 303;
    property ListType: psListType readonly dispid 201;
    property Count: Integer readonly dispid 202;
    property Item[Index: Integer]: IPCsListConnData readonly dispid 0; default;
  end;

// *********************************************************************//
// Interface: IPCsExplorerWindow
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {E3997986-4A75-46F5-9B19-A9C2FE24385B}
// *********************************************************************//
  IPCsExplorerWindow = interface(IDispatch)
    ['{E3997986-4A75-46F5-9B19-A9C2FE24385B}']
    function Get_SubDrawingsFolder: WideString; safecall;
    procedure Set_SubDrawingsFolder(const Value: WideString); safecall;
    function Get_SubDrawingsRootfolder: WideString; safecall;
    procedure Set_SubDrawingsRootfolder(const Value: WideString); safecall;
    property SubDrawingsFolder: WideString read Get_SubDrawingsFolder write Set_SubDrawingsFolder;
    property SubDrawingsRootfolder: WideString read Get_SubDrawingsRootfolder write Set_SubDrawingsRootfolder;
  end;

// *********************************************************************//
// DispIntf:  IPCsExplorerWindowDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {E3997986-4A75-46F5-9B19-A9C2FE24385B}
// *********************************************************************//
  IPCsExplorerWindowDisp = dispinterface
    ['{E3997986-4A75-46F5-9B19-A9C2FE24385B}']
    property SubDrawingsFolder: WideString dispid 201;
    property SubDrawingsRootfolder: WideString dispid 202;
  end;

// *********************************************************************//
// DispIntf:  IPCsExplorerWindowEvents
// Flags:     (0)
// GUID:      {816A2A2D-1F6A-47BD-A36C-FB86AF0EC7AB}
// *********************************************************************//
  IPCsExplorerWindowEvents = dispinterface
    ['{816A2A2D-1F6A-47BD-A36C-FB86AF0EC7AB}']
  end;

// *********************************************************************//
// Interface: IPCsSubDrawing
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0BB8A6D4-2E7D-48D6-9953-CE79FB2AAC33}
// *********************************************************************//
  IPCsSubDrawing = interface(IDispatch)
    ['{0BB8A6D4-2E7D-48D6-9953-CE79FB2AAC33}']
    function Scale(Factor: Double; Pt: TPCsPoint3D): WordBool; safecall;
    function Get_Title: WideString; safecall;
    procedure Set_Title(const Value: WideString); safecall;
    function Get_Valid: WordBool; safecall;
    function Get_Name: WideString; safecall;
    property Title: WideString read Get_Title write Set_Title;
    property Valid: WordBool read Get_Valid;
    property Name: WideString read Get_Name;
  end;

// *********************************************************************//
// DispIntf:  IPCsSubDrawingDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0BB8A6D4-2E7D-48D6-9953-CE79FB2AAC33}
// *********************************************************************//
  IPCsSubDrawingDisp = dispinterface
    ['{0BB8A6D4-2E7D-48D6-9953-CE79FB2AAC33}']
    function Scale(Factor: Double; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant): WordBool; dispid 208;
    property Title: WideString dispid 303;
    property Valid: WordBool readonly dispid 203;
    property Name: WideString readonly dispid 304;
  end;

// *********************************************************************//
// Interface: IPCsListConnData
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {AA0AC284-ACED-46F6-B287-0FCAA2B85BBB}
// *********************************************************************//
  IPCsListConnData = interface(IDispatch)
    ['{AA0AC284-ACED-46F6-B287-0FCAA2B85BBB}']
    procedure GetDataSymbol(DataType: Integer; Options: Integer; out Symbol: IPCsSymbol;
                            out ConnectionIndex: Integer); safecall;
    procedure GetDataLine(DataType: Integer; Options: Integer; out Line: IPCsLine); safecall;
    procedure GetDataInteger(DataType: Integer; Options: Integer; out IntegerValue: Integer); safecall;
  end;

// *********************************************************************//
// DispIntf:  IPCsListConnDataDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {AA0AC284-ACED-46F6-B287-0FCAA2B85BBB}
// *********************************************************************//
  IPCsListConnDataDisp = dispinterface
    ['{AA0AC284-ACED-46F6-B287-0FCAA2B85BBB}']
    procedure GetDataSymbol(DataType: Integer; Options: Integer; out Symbol: IPCsSymbol;
                            out ConnectionIndex: Integer); dispid 203;
    procedure GetDataLine(DataType: Integer; Options: Integer; out Line: IPCsLine); dispid 201;
    procedure GetDataInteger(DataType: Integer; Options: Integer; out IntegerValue: Integer); dispid 202;
  end;

// *********************************************************************//
// Interface: IPCsSelectStatus
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {FA88F8B0-26F1-441C-BBA0-901106CCB211}
// *********************************************************************//
  IPCsSelectStatus = interface(IDispatch)
    ['{FA88F8B0-26F1-441C-BBA0-901106CCB211}']
    function Get_SelectionArea: IPCsSelectionArea; safecall;
    function Get_SelectedSymbol: IPCsSymbol; safecall;
    function Get_SelectedLine: IPCsLine; safecall;
    function Get_SelectedText: IPCsText; safecall;
    function Get_SelectedArc: IPCsArc; safecall;
    function Get_SelectedConnection: IPCsConnection; safecall;
    function MoveSelection(OffsetX: Integer; OffsetY: Integer; OffsetZ: Integer;
                           const Layer: IPCsLayer; const Page: IPCsPage): WordBool; safecall;
    procedure ClearSelection; safecall;
    property SelectionArea: IPCsSelectionArea read Get_SelectionArea;
    property SelectedSymbol: IPCsSymbol read Get_SelectedSymbol;
    property SelectedLine: IPCsLine read Get_SelectedLine;
    property SelectedText: IPCsText read Get_SelectedText;
    property SelectedArc: IPCsArc read Get_SelectedArc;
    property SelectedConnection: IPCsConnection read Get_SelectedConnection;
  end;

// *********************************************************************//
// DispIntf:  IPCsSelectStatusDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {FA88F8B0-26F1-441C-BBA0-901106CCB211}
// *********************************************************************//
  IPCsSelectStatusDisp = dispinterface
    ['{FA88F8B0-26F1-441C-BBA0-901106CCB211}']
    property SelectionArea: IPCsSelectionArea readonly dispid 201;
    property SelectedSymbol: IPCsSymbol readonly dispid 202;
    property SelectedLine: IPCsLine readonly dispid 203;
    property SelectedText: IPCsText readonly dispid 204;
    property SelectedArc: IPCsArc readonly dispid 205;
    property SelectedConnection: IPCsConnection readonly dispid 206;
    function MoveSelection(OffsetX: Integer; OffsetY: Integer; OffsetZ: Integer;
                           const Layer: IPCsLayer; const Page: IPCsPage): WordBool; dispid 207;
    procedure ClearSelection; dispid 208;
  end;

// *********************************************************************//
// DispIntf:  IPCsSelectStatusEvents
// Flags:     (0)
// GUID:      {453EB166-95DD-432B-A435-1A68BF509B8E}
// *********************************************************************//
  IPCsSelectStatusEvents = dispinterface
    ['{453EB166-95DD-432B-A435-1A68BF509B8E}']
  end;

// *********************************************************************//
// Interface: IPCsSelectionArea
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {755A38CB-812B-41CF-9BA6-696F96088873}
// *********************************************************************//
  IPCsSelectionArea = interface(IPCsFigure)
    ['{755A38CB-812B-41CF-9BA6-696F96088873}']
    function Get_Symbols: IPCsSymbols; safecall;
    function Get_Empty: WordBool; safecall;
    property Symbols: IPCsSymbols read Get_Symbols;
    property Empty: WordBool read Get_Empty;
  end;

// *********************************************************************//
// DispIntf:  IPCsSelectionAreaDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {755A38CB-812B-41CF-9BA6-696F96088873}
// *********************************************************************//
  IPCsSelectionAreaDisp = dispinterface
    ['{755A38CB-812B-41CF-9BA6-696F96088873}']
    property Symbols: IPCsSymbols readonly dispid 401;
    property Empty: WordBool readonly dispid 301;
    property Lines: IPCsLines readonly dispid 201;
    property Connections: IPCsLibTypeConnections readonly dispid 202;
    property Arcs: IPCsArcs readonly dispid 203;
    property Texts: IPCsTexts readonly dispid 204;
    property Datafields: IPCsDatafields readonly dispid 205;
    property Title: WideString dispid 303;
    property TitleLC[LCID: Integer]: WideString dispid 206;
    property States: OleVariant readonly dispid 207;
    property SymbolTexts: IPCsSymbolTexts readonly dispid 208;
  end;

// *********************************************************************//
// Interface: IPCsJoins
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {13FA989D-485B-4BE2-9062-148F69DA8D33}
// *********************************************************************//
  IPCsJoins = interface(IDispatch)
    ['{13FA989D-485B-4BE2-9062-148F69DA8D33}']
    function Get_Count: Integer; safecall;
    function Get__NewEnum: IEnumVARIANT; safecall;
    function Get_Item(Index: Integer): IPCsJoin; safecall;
    property Count: Integer read Get_Count;
    property _NewEnum: IEnumVARIANT read Get__NewEnum;
    property Item[Index: Integer]: IPCsJoin read Get_Item; default;
  end;

// *********************************************************************//
// DispIntf:  IPCsJoinsDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {13FA989D-485B-4BE2-9062-148F69DA8D33}
// *********************************************************************//
  IPCsJoinsDisp = dispinterface
    ['{13FA989D-485B-4BE2-9062-148F69DA8D33}']
    property Count: Integer readonly dispid 202;
    property _NewEnum: IEnumVARIANT readonly dispid -4;
    property Item[Index: Integer]: IPCsJoin readonly dispid 0; default;
  end;

// *********************************************************************//
// Interface: IPCsJoin
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {5B281FCC-D09C-4C0C-935C-EFE272D57E0F}
// *********************************************************************//
  IPCsJoin = interface(IDispatch)
    ['{5B281FCC-D09C-4C0C-935C-EFE272D57E0F}']
    function Get_Item(Index: Integer): IPCsJoinPoint; safecall;
    function Get_Count: Integer; safecall;
    function Get__NewEnum: IEnumVARIANT; safecall;
    property Item[Index: Integer]: IPCsJoinPoint read Get_Item; default;
    property Count: Integer read Get_Count;
    property _NewEnum: IEnumVARIANT read Get__NewEnum;
  end;

// *********************************************************************//
// DispIntf:  IPCsJoinDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {5B281FCC-D09C-4C0C-935C-EFE272D57E0F}
// *********************************************************************//
  IPCsJoinDisp = dispinterface
    ['{5B281FCC-D09C-4C0C-935C-EFE272D57E0F}']
    property Item[Index: Integer]: IPCsJoinPoint readonly dispid 0; default;
    property Count: Integer readonly dispid 201;
    property _NewEnum: IEnumVARIANT readonly dispid -4;
  end;

// *********************************************************************//
// Interface: IPCsDrawingPreferences
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {66BFC81B-1CCA-4E53-8C15-51A80752534B}
// *********************************************************************//
  IPCsDrawingPreferences = interface(IDispatch)
    ['{66BFC81B-1CCA-4E53-8C15-51A80752534B}']
    function Get_SymbolNameFormat: WideString; safecall;
    procedure Set_SymbolNameFormat(const Value: WideString); safecall;
    function Get_ReferenceDesignationWithLinebreak: WordBool; safecall;
    procedure Set_ReferenceDesignationWithLinebreak(Value: WordBool); safecall;
    function Get_MountingCorrectDrawing: WordBool; safecall;
    procedure Set_MountingCorrectDrawing(Value: WordBool); safecall;
    function Get_ZipComment: WideString; safecall;
    function Get_VisibleTexts(TextType: PsTextType): WordBool; safecall;
    procedure Set_VisibleTexts(TextType: PsTextType; Value: WordBool); safecall;
    function Get_DiaNormalSnap: Integer; safecall;
    procedure Set_DiaNormalSnap(Value: Integer); safecall;
    function Get_DiaFineSnap: Integer; safecall;
    procedure Set_DiaFineSnap(Value: Integer); safecall;
    function Get_GrpNormalSnap: Integer; safecall;
    procedure Set_GrpNormalSnap(Value: Integer); safecall;
    function Get_GrpFineSnap: Integer; safecall;
    procedure Set_GrpFineSnap(Value: Integer); safecall;
    property SymbolNameFormat: WideString read Get_SymbolNameFormat write Set_SymbolNameFormat;
    property ReferenceDesignationWithLinebreak: WordBool read Get_ReferenceDesignationWithLinebreak write Set_ReferenceDesignationWithLinebreak;
    property MountingCorrectDrawing: WordBool read Get_MountingCorrectDrawing write Set_MountingCorrectDrawing;
    property ZipComment: WideString read Get_ZipComment;
    property VisibleTexts[TextType: PsTextType]: WordBool read Get_VisibleTexts write Set_VisibleTexts;
    property DiaNormalSnap: Integer read Get_DiaNormalSnap write Set_DiaNormalSnap;
    property DiaFineSnap: Integer read Get_DiaFineSnap write Set_DiaFineSnap;
    property GrpNormalSnap: Integer read Get_GrpNormalSnap write Set_GrpNormalSnap;
    property GrpFineSnap: Integer read Get_GrpFineSnap write Set_GrpFineSnap;
  end;

// *********************************************************************//
// DispIntf:  IPCsDrawingPreferencesDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {66BFC81B-1CCA-4E53-8C15-51A80752534B}
// *********************************************************************//
  IPCsDrawingPreferencesDisp = dispinterface
    ['{66BFC81B-1CCA-4E53-8C15-51A80752534B}']
    property SymbolNameFormat: WideString dispid 201;
    property ReferenceDesignationWithLinebreak: WordBool dispid 202;
    property MountingCorrectDrawing: WordBool dispid 203;
    property ZipComment: WideString readonly dispid 204;
    property VisibleTexts[TextType: PsTextType]: WordBool dispid 205;
    property DiaNormalSnap: Integer dispid 206;
    property DiaFineSnap: Integer dispid 207;
    property GrpNormalSnap: Integer dispid 208;
    property GrpFineSnap: Integer dispid 209;
  end;

// *********************************************************************//
// Interface: IPCsReference
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {3077DBD0-D206-4478-A21F-E709AC5D1EDE}
// *********************************************************************//
  IPCsReference = interface(IDispatch)
    ['{3077DBD0-D206-4478-A21F-E709AC5D1EDE}']
    function Get_ReferenceObj: IPCsObject; safecall;
    function Get_Pretext: WideString; safecall;
    procedure Set_Pretext(const Value: WideString); safecall;
    function Get_Position: TPCsPoint3D; safecall;
    procedure Set_Position(Value: TPCsPoint3D); safecall;
    function Get_Value: WideString; safecall;
    function Get_ShowAllReferences: WordBool; safecall;
    procedure Set_ShowAllReferences(Value: WordBool); safecall;
    function Get_InColumn: WordBool; safecall;
    procedure Set_InColumn(Value: WordBool); safecall;
    property ReferenceObj: IPCsObject read Get_ReferenceObj;
    property Pretext: WideString read Get_Pretext write Set_Pretext;
    property Position: TPCsPoint3D read Get_Position write Set_Position;
    property Value: WideString read Get_Value;
    property ShowAllReferences: WordBool read Get_ShowAllReferences write Set_ShowAllReferences;
    property InColumn: WordBool read Get_InColumn write Set_InColumn;
  end;

// *********************************************************************//
// DispIntf:  IPCsReferenceDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {3077DBD0-D206-4478-A21F-E709AC5D1EDE}
// *********************************************************************//
  IPCsReferenceDisp = dispinterface
    ['{3077DBD0-D206-4478-A21F-E709AC5D1EDE}']
    property ReferenceObj: IPCsObject readonly dispid 201;
    property Pretext: WideString dispid 202;
    property Position: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant dispid 203;
    property Value: WideString readonly dispid 204;
    property ShowAllReferences: WordBool dispid 205;
    property InColumn: WordBool dispid 206;
  end;

// *********************************************************************//
// Interface: IPCsReferenceFrame
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {7FED85F3-51A8-4B62-80BA-6F615B6206FF}
// *********************************************************************//
  IPCsReferenceFrame = interface(IPCsLine)
    ['{7FED85F3-51A8-4B62-80BA-6F615B6206FF}']
    function Get_NameText: IPCsNameText; safecall;
    function Get_DisplayText: WideString; safecall;
    property NameText: IPCsNameText read Get_NameText;
    property DisplayText: WideString read Get_DisplayText;
  end;

// *********************************************************************//
// DispIntf:  IPCsReferenceFrameDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {7FED85F3-51A8-4B62-80BA-6F615B6206FF}
// *********************************************************************//
  IPCsReferenceFrameDisp = dispinterface
    ['{7FED85F3-51A8-4B62-80BA-6F615B6206FF}']
    property NameText: IPCsNameText readonly dispid 314;
    property DisplayText: WideString readonly dispid 320;
    property Points: PCsPoints readonly dispid 401;
    procedure AddPoint(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 402;
    property IsElectrical: WordBool dispid 301;
    property LineStyle: PsLineStyle dispid 302;
    property Color: PsPenColor dispid 303;
    property PenWidth: Integer dispid 304;
    property Quantity: Double readonly dispid 305;
    property IsFilled: WordBool dispid 306;
    property IsJumperingLink: WordBool dispid 308;
    property IsCableWire: WordBool dispid 307;
    property IsPartLine: WordBool readonly dispid 310;
    property Extended: WordBool dispid 311;
    property IsOpen: WordBool dispid 312;
    property IsSignalsJoined: WordBool readonly dispid 313;
    function AddBendingPoint(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; BendFactor: Double): WordBool; dispid 315;
    procedure ConnectTo(const JoinPoint: IPCsJoinPoint); dispid 316;
    property Pattern: Integer dispid 317;
    property MultilineDistance: Integer dispid 318;
    property Layer: IPCsLayer dispid 319;
    property IsCurvedLine: WordBool dispid 309;
    property Visible: WordBool dispid 201;
    procedure AddPointXYZ(X: Integer; Y: Integer; Z: Integer); dispid 202;
    procedure Delete; dispid 203;
    property LineTexts: IPCsLineTexts readonly dispid 204;
    property Handle: Integer readonly dispid 205;
    property State: Integer dispid 250;
    property ObjectType: PsPCsObjectType readonly dispid 450;
    procedure InitLineTexts; dispid 206;
    function MoveRelative(AOffset: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Figure: IPCsFigure): WordBool; dispid 207;
    function MoveRelativeXYZ(OffsetX: Integer; OffsetY: Integer; OffsetZ: Integer;
                             const Figure: IPCsFigure): WordBool; dispid 208;
    function ChangeDrawingOrder(AAction: psChangeDrawingOrder): WordBool; dispid 209;
  end;

// *********************************************************************//
// Interface: IPCsSymbol3
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {68FA7365-D511-43B4-B220-B01EA114EA8A}
// *********************************************************************//
  IPCsSymbol3 = interface(IPCsSymbol2)
    ['{68FA7365-D511-43B4-B220-B01EA114EA8A}']
    function Get_IsFeedThroughTerminal: WordBool; safecall;
    function Get_IsTerminal: WordBool; safecall;
    procedure Set_IsTerminal(Value: WordBool); safecall;
    function IsSameComponent(const ASymbol: IPCsSymbol; Options: psIsSameComponentOptions): WordBool; safecall;
    function IsIdenticalTo(const ASymbol: IPCsSymbol): WordBool; safecall;
    function Get_LayerIndex: Integer; safecall;
    procedure Set_LayerIndex(Value: Integer); safecall;
    function Get_IsElectrical: WordBool; safecall;
    procedure Set_IsElectrical(Value: WordBool); safecall;
    function Get_IsMechanical: WordBool; safecall;
    procedure Set_IsMechanical(Value: WordBool); safecall;
    function Get_IncludeInMechanicalLoad: WordBool; safecall;
    procedure Set_IncludeInMechanicalLoad(Value: WordBool); safecall;
    function Get_IsProtected: WordBool; safecall;
    procedure Set_IsProtected(Value: WordBool); safecall;
    function IsSamePin(ConnectionIndex: Integer; const OtherSymbol: IPCsSymbol;
                       OtherConnectionIndex: Integer; Options: psIsSameComponentOptions): WordBool; safecall;
    function IsConnectionPair(ConnectionIndex1: Integer; ConnectionIndex2: Integer): WordBool; safecall;
    function HasSameNameAs(const ASymbol: IPCsSymbol): WordBool; safecall;
    function IsMetaconnected(ConnectionIndex: Integer; const OtherSymbol: IPCsSymbol;
                             OtherConnectionIndex: Integer; Options: Integer): WordBool; safecall;
    procedure SetQuantity(Quantity: Double); safecall;
    property IsFeedThroughTerminal: WordBool read Get_IsFeedThroughTerminal;
    property IsTerminal: WordBool read Get_IsTerminal write Set_IsTerminal;
    property LayerIndex: Integer read Get_LayerIndex write Set_LayerIndex;
    property IsElectrical: WordBool read Get_IsElectrical write Set_IsElectrical;
    property IsMechanical: WordBool read Get_IsMechanical write Set_IsMechanical;
    property IncludeInMechanicalLoad: WordBool read Get_IncludeInMechanicalLoad write Set_IncludeInMechanicalLoad;
    property IsProtected: WordBool read Get_IsProtected write Set_IsProtected;
  end;

// *********************************************************************//
// DispIntf:  IPCsSymbol3Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {68FA7365-D511-43B4-B220-B01EA114EA8A}
// *********************************************************************//
  IPCsSymbol3Disp = dispinterface
    ['{68FA7365-D511-43B4-B220-B01EA114EA8A}']
    property IsFeedThroughTerminal: WordBool readonly dispid 501;
    property IsTerminal: WordBool dispid 502;
    function IsSameComponent(const ASymbol: IPCsSymbol; Options: psIsSameComponentOptions): WordBool; dispid 503;
    function IsIdenticalTo(const ASymbol: IPCsSymbol): WordBool; dispid 504;
    property LayerIndex: Integer dispid 505;
    property IsElectrical: WordBool dispid 506;
    property IsMechanical: WordBool dispid 507;
    property IncludeInMechanicalLoad: WordBool dispid 508;
    property IsProtected: WordBool dispid 509;
    function IsSamePin(ConnectionIndex: Integer; const OtherSymbol: IPCsSymbol;
                       OtherConnectionIndex: Integer; Options: psIsSameComponentOptions): WordBool; dispid 510;
    function IsConnectionPair(ConnectionIndex1: Integer; ConnectionIndex2: Integer): WordBool; dispid 511;
    function HasSameNameAs(const ASymbol: IPCsSymbol): WordBool; dispid 512;
    function IsMetaconnected(ConnectionIndex: Integer; const OtherSymbol: IPCsSymbol;
                             OtherConnectionIndex: Integer; Options: Integer): WordBool; dispid 513;
    procedure SetQuantity(Quantity: Double); dispid 514;
    property SymbolData[Index: Integer]: WideString dispid 451;
    property SymbolDataCount: Integer readonly dispid 452;
    function Equals(const Obj: IPCsBase): WordBool; dispid 401;
    function IsErased: WordBool; dispid 403;
    property LibType: IPCsLibType readonly dispid 402;
    property IsVMirrored: WordBool dispid 405;
    property IsHMirrored: WordBool dispid 406;
    property SymbolTexts: IPCsSymbolTexts readonly dispid 407;
    property Page: IPCsPage readonly dispid 408;
    property Layer: IPCsLayer dispid 409;
    property FullName: WideString dispid 302;
    function GetRect(AType: PsSymbolRectType): {NOT_OLEAUTO(TPCsRect)}OleVariant; dispid 303;
    property Connections: IPCsConnections readonly dispid 304;
    property Rotation: Integer dispid 305;
    property ScaleFactor: Double dispid 306;
    function Explode(const ADestination: IPCsFigure): WordBool; dispid 307;
    property SymbolType1: psSymbolType dispid 301;
    property SymbolType2: psSymbolType dispid 308;
    procedure ApplicationDataWriteString(const ASection: WideString; const AIdent: WideString;
                                         const AValue: WideString); dispid 309;
    function ApplicationDataReadString(const ASection: WideString; const AIdent: WideString;
                                       const ADefault: WideString): WideString; dispid 310;
    property Datafields: IPCsDatafields readonly dispid 311;
    procedure Delete; dispid 312;
    property ObjectType: PsPCsObjectType readonly dispid 450;
    property ComponentGroupNo: Integer dispid 313;
    property ComponentPosNo: Integer dispid 314;
    property Reference: IPCsReference readonly dispid 315;
    property WithReference: WordBool dispid 316;
    function UpdateFromDatabase(Reserved: Integer): WordBool; dispid 317;
    property Visible: WordBool dispid 201;
    property Position: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant dispid 202;
    property LStatus: Integer dispid 203;
    function Move(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Figure: IPCsFigure): WordBool; dispid 204;
    procedure Rotate(DeltaAngle: Integer; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 205;
    procedure Scale(Factor: Single; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 208;
    property Handle: LongWord readonly dispid 206;
    procedure GetPosition(out X: Integer; out Y: Integer; out Z: Integer); dispid 207;
    procedure SetPosition(X: Integer; Y: Integer; Z: Integer); dispid 209;
    property State: Integer dispid 250;
    function ChangeDrawingOrder(AAction: psChangeDrawingOrder): WordBool; dispid 210;
  end;

// *********************************************************************//
// Interface: IPCsLine3
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {41776201-3CAF-4CB0-A10F-D9DE66846CBF}
// *********************************************************************//
  IPCsLine3 = interface(IPCsLine2)
    ['{41776201-3CAF-4CB0-A10F-D9DE66846CBF}']
    procedure UniteJoinedLines; safecall;
    procedure RemoveExtendedSegments; safecall;
    procedure RemoveZeroVectors; safecall;
    function CanUniteWith(const ALine: IPCsLine): WordBool; safecall;
    function UniteWith(const ALine: IPCsLine): WordBool; safecall;
  end;

// *********************************************************************//
// DispIntf:  IPCsLine3Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {41776201-3CAF-4CB0-A10F-D9DE66846CBF}
// *********************************************************************//
  IPCsLine3Disp = dispinterface
    ['{41776201-3CAF-4CB0-A10F-D9DE66846CBF}']
    procedure UniteJoinedLines; dispid 403;
    procedure RemoveExtendedSegments; dispid 404;
    procedure RemoveZeroVectors; dispid 405;
    function CanUniteWith(const ALine: IPCsLine): WordBool; dispid 406;
    function UniteWith(const ALine: IPCsLine): WordBool; dispid 407;
    function MoveLinePoint(PointIndex: Integer; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant;
                           Options: Integer): WordBool; dispid 314;
    function MoveLinePointXYZ(PointIndex: Integer; X: Integer; Y: Integer; Z: Integer;
                              Options: Integer): WordBool; dispid 320;
    property Application: IPCsApplication readonly dispid 321;
    function Equals(const Obj: IPCsBase): WordBool; dispid 322;
    function IsErased: WordBool; dispid 323;
    property Points: PCsPoints readonly dispid 401;
    procedure AddPoint(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 402;
    property IsElectrical: WordBool dispid 301;
    property LineStyle: PsLineStyle dispid 302;
    property Color: PsPenColor dispid 303;
    property PenWidth: Integer dispid 304;
    property Quantity: Double readonly dispid 305;
    property IsFilled: WordBool dispid 306;
    property IsJumperingLink: WordBool dispid 308;
    property IsCableWire: WordBool dispid 307;
    property IsPartLine: WordBool readonly dispid 310;
    property Extended: WordBool dispid 311;
    property IsOpen: WordBool dispid 312;
    property IsSignalsJoined: WordBool readonly dispid 313;
    function AddBendingPoint(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; BendFactor: Double): WordBool; dispid 315;
    procedure ConnectTo(const JoinPoint: IPCsJoinPoint); dispid 316;
    property Pattern: Integer dispid 317;
    property MultilineDistance: Integer dispid 318;
    property Layer: IPCsLayer dispid 319;
    property IsCurvedLine: WordBool dispid 309;
    property Visible: WordBool dispid 201;
    procedure AddPointXYZ(X: Integer; Y: Integer; Z: Integer); dispid 202;
    procedure Delete; dispid 203;
    property LineTexts: IPCsLineTexts readonly dispid 204;
    property Handle: Integer readonly dispid 205;
    property State: Integer dispid 250;
    property ObjectType: PsPCsObjectType readonly dispid 450;
    procedure InitLineTexts; dispid 206;
    function MoveRelative(AOffset: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Figure: IPCsFigure): WordBool; dispid 207;
    function MoveRelativeXYZ(OffsetX: Integer; OffsetY: Integer; OffsetZ: Integer;
                             const Figure: IPCsFigure): WordBool; dispid 208;
    function ChangeDrawingOrder(AAction: psChangeDrawingOrder): WordBool; dispid 209;
  end;

// *********************************************************************//
// Interface: IPCsPage3
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {2A6D4CCB-9F5E-4B79-A285-BF1A375320D6}
// *********************************************************************//
  IPCsPage3 = interface(IDispatch)
    ['{2A6D4CCB-9F5E-4B79-A285-BF1A375320D6}']
    procedure GetDrawingArea(out X1: Integer; out Y1: Integer; out X2: Integer; out Y2: Integer); safecall;
    function Get_GridSize: Integer; safecall;
    procedure Set_GridSize(Value: Integer); safecall;
    function Get_HighlightedObjectsCount(Kind: psHighlighting;
                                         Options: psIterateDrawingObjectsOptions): Integer; safecall;
    function Get_LastModifiedTime: Double; safecall;
    procedure Set_LastModifiedTime(Value: Double); safecall;
    function PlaceSymbolXYZ(X: Integer; Y: Integer; Z: Integer; const Filename: WideString): IPCsSymbol3; safecall;
    property GridSize: Integer read Get_GridSize write Set_GridSize;
    property HighlightedObjectsCount[Kind: psHighlighting; Options: psIterateDrawingObjectsOptions]: Integer read Get_HighlightedObjectsCount;
    property LastModifiedTime: Double read Get_LastModifiedTime write Set_LastModifiedTime;
  end;

// *********************************************************************//
// DispIntf:  IPCsPage3Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {2A6D4CCB-9F5E-4B79-A285-BF1A375320D6}
// *********************************************************************//
  IPCsPage3Disp = dispinterface
    ['{2A6D4CCB-9F5E-4B79-A285-BF1A375320D6}']
    procedure GetDrawingArea(out X1: Integer; out Y1: Integer; out X2: Integer; out Y2: Integer); dispid 502;
    property GridSize: Integer dispid 504;
    property HighlightedObjectsCount[Kind: psHighlighting; Options: psIterateDrawingObjectsOptions]: Integer readonly dispid 202;
    property LastModifiedTime: Double dispid 201;
    function PlaceSymbolXYZ(X: Integer; Y: Integer; Z: Integer; const Filename: WideString): IPCsSymbol3; dispid 203;
  end;

// *********************************************************************//
// Interface: IPCsDrawing3
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {24AB84D8-EB62-4160-8238-C5FD95598D4F}
// *********************************************************************//
  IPCsDrawing3 = interface(IPCsDrawing2)
    ['{24AB84D8-EB62-4160-8238-C5FD95598D4F}']
    function Get_HighlightedObjectsCount(Kind: psHighlighting;
                                         Options: psIterateDrawingObjectsOptions): Integer; safecall;
    function GetSymbolFromHandleUnsafe(APageIndex: Integer; AHandle: Integer): IPCsSymbol2; safecall;
    function GetLineFromHandleUnsafe(APageIndex: Integer; AHandle: Integer): IPCsLine4; safecall;
    function GetTextFromHandleUnsafe(APageIndex: Integer; AHandle: Integer): IPCsText; safecall;
    procedure AddEventListenerWM(const AType: WideString; AWindowHandle: TWindowHandle;
                                 AMessage: Integer); safecall;
    procedure RemoveEventListenersWM(AWindowHandle: TWindowHandle); safecall;
    procedure ResetHighlighting; safecall;
    procedure UniteJoinedEqualLines; safecall;
    procedure CheckJoins(ARepare: WordBool); safecall;
    procedure RemoveZeroVectors; safecall;
    property HighlightedObjectsCount[Kind: psHighlighting; Options: psIterateDrawingObjectsOptions]: Integer read Get_HighlightedObjectsCount;
  end;

// *********************************************************************//
// DispIntf:  IPCsDrawing3Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {24AB84D8-EB62-4160-8238-C5FD95598D4F}
// *********************************************************************//
  IPCsDrawing3Disp = dispinterface
    ['{24AB84D8-EB62-4160-8238-C5FD95598D4F}']
    property HighlightedObjectsCount[Kind: psHighlighting; Options: psIterateDrawingObjectsOptions]: Integer readonly dispid 401;
    function GetSymbolFromHandleUnsafe(APageIndex: Integer; AHandle: Integer): IPCsSymbol2; dispid 402;
    function GetLineFromHandleUnsafe(APageIndex: Integer; AHandle: Integer): IPCsLine4; dispid 403;
    function GetTextFromHandleUnsafe(APageIndex: Integer; AHandle: Integer): IPCsText; dispid 404;
    procedure AddEventListenerWM(const AType: WideString; AWindowHandle: TWindowHandle;
                                 AMessage: Integer); dispid 405;
    procedure RemoveEventListenersWM(AWindowHandle: TWindowHandle); dispid 406;
    procedure ResetHighlighting; dispid 407;
    procedure UniteJoinedEqualLines; dispid 408;
    procedure CheckJoins(ARepare: WordBool); dispid 409;
    procedure RemoveZeroVectors; dispid 410;
    function GetUsedLibtypes: OleVariant; dispid 301;
    function Equals(const Obj: IPCsBase): WordBool; dispid 302;
    function IsErased: WordBool; dispid 303;
    property Logos[Index: Integer]: IPCsExternalObject dispid 304;
    function LoadLogo(Index: Integer; const Filename: WideString): WordBool; dispid 305;
    property LogosCount: Integer readonly dispid 306;
    function ReplaceSymbol(const SymbolToReplace: IPCsSymbol;
                           const SymbolToReplaceWith: IPCsSymbol; Options: Integer): IPCsSymbol; dispid 307;
    property Pages: IPCsPages readonly dispid 201;
    property LibTypes: IPCsLibTypes readonly dispid 202;
    property Layers: IPCsLayers readonly dispid 205;
    property Application: IPCsApplication readonly dispid 206;
    property Parent: IPCsDocument readonly dispid 207;
    property Title: WideString dispid 203;
    function GetSymbolFromHandle(PageNo: Integer; SymbolHandle: Integer): IPCsSymbol; dispid 204;
    property ActiveLayer: IPCsLayer dispid 208;
    property ProjectData[Index: Integer]: WideString dispid 209;
    property ProjectDataDefs: OleVariant dispid 210;
    property PageDataDefs: OleVariant dispid 211;
    property SymbolDataDefs: OleVariant dispid 214;
    property ReferenceDesignations: IPCsReferenceDesignations readonly dispid 212;
    function SetProjectData(const FieldName: WideString; const Value: WideString): Integer; dispid 213;
    property Lists: IPCsLists readonly dispid 215;
    procedure OptimizeLibtypes; dispid 216;
    function LoadSubDrawing(const Filename: WideString): IPCsSubDrawing; dispid 217;
    function GetLineFromHandle(PageNo: Integer; LineHandle: Integer): IPCsLine; dispid 218;
    function GetTextFromHandle(PageNo: Integer; TextHandle: Integer): IPCsText; dispid 219;
    property SystemVariable[const Variable: WideString]: OleVariant dispid 220;
    property Preferences: IPCsDrawingPreferences readonly dispid 221;
    procedure SynchronizePLCData(const ParamStr: WideString); dispid 222;
    procedure ApplyIntelligentInvisibility(const ParamStr: WideString); dispid 223;
    function UpdateFromDatabase(Reserved: Integer): WordBool; dispid 224;
    property Document: IPCsDocument readonly dispid 225;
  end;

// *********************************************************************//
// Interface: IPCsGlobal3
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {212D442E-0E51-45DF-809F-DAB5397B80D0}
// *********************************************************************//
  IPCsGlobal3 = interface(IPCsGlobal2)
    ['{212D442E-0E51-45DF-809F-DAB5397B80D0}']
    function BuildReferenceDesignationStringEx(FunctionalAspect: OleVariant;
                                               LocationAspect: OleVariant;
                                               ProductAspect: OleVariant;
                                               const ComponentName: WideString;
                                               ObjectType: PsPCsObjectType;
                                               const Format: IPCsRefIDFormat): WideString; safecall;
  end;

// *********************************************************************//
// DispIntf:  IPCsGlobal3Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {212D442E-0E51-45DF-809F-DAB5397B80D0}
// *********************************************************************//
  IPCsGlobal3Disp = dispinterface
    ['{212D442E-0E51-45DF-809F-DAB5397B80D0}']
    function BuildReferenceDesignationStringEx(FunctionalAspect: OleVariant;
                                               LocationAspect: OleVariant;
                                               ProductAspect: OleVariant;
                                               const ComponentName: WideString;
                                               ObjectType: PsPCsObjectType;
                                               const Format: IPCsRefIDFormat): WideString; dispid 304;
    function BuildReferenceDesignationString(const FunctionalAspect: WideString;
                                             const LocationAspect: WideString;
                                             const ProductAspect: WideString;
                                             const ComponentName: WideString;
                                             const FormatString: WideString): WideString; dispid 301;
  end;

// *********************************************************************//
// Interface: IPCsDrawingPreferences2
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {BA0AFD46-55FF-48BA-87F2-5FC856AE81C7}
// *********************************************************************//
  IPCsDrawingPreferences2 = interface(IPCsDrawingPreferences)
    ['{BA0AFD46-55FF-48BA-87F2-5FC856AE81C7}']
    function Get_ReferenceDesignationFormat1: WideString; safecall;
    procedure Set_ReferenceDesignationFormat1(const Value: WideString); safecall;
    function Get_ReferenceDesignationFormat2: WideString; safecall;
    procedure Set_ReferenceDesignationFormat2(const Value: WideString); safecall;
    function Get_Application: IPCsApplication2; safecall;
    property ReferenceDesignationFormat1: WideString read Get_ReferenceDesignationFormat1 write Set_ReferenceDesignationFormat1;
    property ReferenceDesignationFormat2: WideString read Get_ReferenceDesignationFormat2 write Set_ReferenceDesignationFormat2;
    property Application: IPCsApplication2 read Get_Application;
  end;

// *********************************************************************//
// DispIntf:  IPCsDrawingPreferences2Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {BA0AFD46-55FF-48BA-87F2-5FC856AE81C7}
// *********************************************************************//
  IPCsDrawingPreferences2Disp = dispinterface
    ['{BA0AFD46-55FF-48BA-87F2-5FC856AE81C7}']
    property ReferenceDesignationFormat1: WideString dispid 303;
    property ReferenceDesignationFormat2: WideString dispid 304;
    property Application: IPCsApplication2 readonly dispid 301;
    property SymbolNameFormat: WideString dispid 201;
    property ReferenceDesignationWithLinebreak: WordBool dispid 202;
    property MountingCorrectDrawing: WordBool dispid 203;
    property ZipComment: WideString readonly dispid 204;
    property VisibleTexts[TextType: PsTextType]: WordBool dispid 205;
    property DiaNormalSnap: Integer dispid 206;
    property DiaFineSnap: Integer dispid 207;
    property GrpNormalSnap: Integer dispid 208;
    property GrpFineSnap: Integer dispid 209;
  end;

// *********************************************************************//
// Interface: IPCsNameText3
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {9305B2C0-5267-4A89-8AA5-06A5842C07C1}
// *********************************************************************//
  IPCsNameText3 = interface(IPCsNameText2)
    ['{9305B2C0-5267-4A89-8AA5-06A5842C07C1}']
    function Get_RefDesString(Aspect: psReferenceDesignationAspect): WideString; safecall;
    procedure Set_RefDesString(Aspect: psReferenceDesignationAspect; const Value: WideString); safecall;
    function Get_ReferenceDesignation: IPCsObjectRefID; safecall;
    property RefDesString[Aspect: psReferenceDesignationAspect]: WideString read Get_RefDesString write Set_RefDesString;
    property ReferenceDesignation: IPCsObjectRefID read Get_ReferenceDesignation;
  end;

// *********************************************************************//
// DispIntf:  IPCsNameText3Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {9305B2C0-5267-4A89-8AA5-06A5842C07C1}
// *********************************************************************//
  IPCsNameText3Disp = dispinterface
    ['{9305B2C0-5267-4A89-8AA5-06A5842C07C1}']
    property RefDesString[Aspect: psReferenceDesignationAspect]: WideString dispid 600;
    property ReferenceDesignation: IPCsObjectRefID readonly dispid 601;
    property DisplayText2: WideString readonly dispid 501;
    function Equals(const Obj: IPCsBase): WordBool; dispid 502;
    property Application: IPCsApplication2 readonly dispid 503;
    function IsErased: WordBool; dispid 504;
    property RefDesFunction: WideString dispid 409;
    property RefDesLocation: WideString dispid 410;
    property RefDesOptions: psSymbolRefDesOption dispid 411;
    property Value: WideString dispid 401;
    property Origin: PsTextOrigin dispid 402;
    property Font: IPCsTextFont readonly dispid 403;
    property TextType: PsTextOrigin readonly dispid 404;
    property FieldWidth: Integer dispid 405;
    property WrapText: WordBool dispid 406;
    property LinkID: Integer dispid 407;
    property PositionExt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant dispid 408;
    property Color: PsPenColor dispid 301;
    property Layer: IPCsLayer dispid 302;
    property DisplayText: WideString readonly dispid 303;
    property LineCount: Integer readonly dispid 304;
    property Rotation: Integer dispid 305;
    property IsBlank: WordBool readonly dispid 306;
    function GetRect(AFlags: psTextRectFlags): {NOT_OLEAUTO(TPCsRect)}OleVariant; dispid 307;
    procedure Delete; dispid 308;
    property ObjectType: PsPCsObjectType readonly dispid 450;
    property NotTranslated: WordBool dispid 309;
    property Visible: WordBool dispid 201;
    property Position: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant dispid 202;
    property LStatus: Integer dispid 203;
    function Move(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Figure: IPCsFigure): WordBool; dispid 204;
    procedure Rotate(DeltaAngle: Integer; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 205;
    procedure Scale(Factor: Single; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 208;
    property Handle: LongWord readonly dispid 206;
    procedure GetPosition(out X: Integer; out Y: Integer; out Z: Integer); dispid 207;
    procedure SetPosition(X: Integer; Y: Integer; Z: Integer); dispid 209;
    property State: Integer dispid 250;
    function ChangeDrawingOrder(AAction: psChangeDrawingOrder): WordBool; dispid 210;
  end;

// *********************************************************************//
// Interface: IPCsSubDrawing2
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {6B8CCDD1-805E-43A9-8B54-0444DA25F734}
// *********************************************************************//
  IPCsSubDrawing2 = interface(IPCsSubDrawing)
    ['{6B8CCDD1-805E-43A9-8B54-0444DA25F734}']
    function Get_TemplateData: IPCsTemplateData; safecall;
    function Equals(const Obj: IPCsBase): WordBool; safecall;
    function Get_Application: IPCsApplication2; safecall;
    function IsErased: WordBool; safecall;
    property TemplateData: IPCsTemplateData read Get_TemplateData;
    property Application: IPCsApplication2 read Get_Application;
  end;

// *********************************************************************//
// DispIntf:  IPCsSubDrawing2Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {6B8CCDD1-805E-43A9-8B54-0444DA25F734}
// *********************************************************************//
  IPCsSubDrawing2Disp = dispinterface
    ['{6B8CCDD1-805E-43A9-8B54-0444DA25F734}']
    property TemplateData: IPCsTemplateData readonly dispid 301;
    function Equals(const Obj: IPCsBase): WordBool; dispid 302;
    property Application: IPCsApplication2 readonly dispid 305;
    function IsErased: WordBool; dispid 306;
    function Scale(Factor: Double; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant): WordBool; dispid 208;
    property Title: WideString dispid 303;
    property Valid: WordBool readonly dispid 203;
    property Name: WideString readonly dispid 304;
  end;

// *********************************************************************//
// Interface: IPCsTemplateData
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {09646609-A0D9-49A5-95B3-483E77F7E479}
// *********************************************************************//
  IPCsTemplateData = interface(IDispatch)
    ['{09646609-A0D9-49A5-95B3-483E77F7E479}']
    function Get_ModelTitles: OleVariant; safecall;
    property ModelTitles: OleVariant read Get_ModelTitles;
  end;

// *********************************************************************//
// DispIntf:  IPCsTemplateDataDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {09646609-A0D9-49A5-95B3-483E77F7E479}
// *********************************************************************//
  IPCsTemplateDataDisp = dispinterface
    ['{09646609-A0D9-49A5-95B3-483E77F7E479}']
    property ModelTitles: OleVariant readonly dispid 201;
  end;

// *********************************************************************//
// Interface: IPCsExternalObject
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {E51DE5D6-F827-476C-A50A-08BDA31F3257}
// *********************************************************************//
  IPCsExternalObject = interface(IPCsStatusPoint)
    ['{E51DE5D6-F827-476C-A50A-08BDA31F3257}']
    function Get_IsPicture: WordBool; safecall;
    function Get_IsOleObject: WordBool; safecall;
    function Get_Width: Integer; safecall;
    procedure Set_Width(Value: Integer); safecall;
    function Get_Height: Integer; safecall;
    procedure Set_Height(Value: Integer); safecall;
    procedure SaveToFile(const Filename: WideString; Options: Integer); safecall;
    procedure SaveToStream(const Stream: IUnknown; Options: Integer); safecall;
    procedure LoadFromStream(const Stream: IUnknown); safecall;
    function Get_Filename: WideString; safecall;
    procedure Set_Filename(const Value: WideString); safecall;
    function CustomCommand(const Command: WideString; Params: OleVariant; out Results: OleVariant): WordBool; safecall;
    function IsErased: WordBool; safecall;
    procedure AssignTo(const Obj: IPCsBase); safecall;
    function Equals(const Obj: IPCsBase): WordBool; safecall;
    function Get_IsValid: WordBool; safecall;
    function Get_IsLinked: WordBool; safecall;
    property IsPicture: WordBool read Get_IsPicture;
    property IsOleObject: WordBool read Get_IsOleObject;
    property Width: Integer read Get_Width write Set_Width;
    property Height: Integer read Get_Height write Set_Height;
    property Filename: WideString read Get_Filename write Set_Filename;
    property IsValid: WordBool read Get_IsValid;
    property IsLinked: WordBool read Get_IsLinked;
  end;

// *********************************************************************//
// DispIntf:  IPCsExternalObjectDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {E51DE5D6-F827-476C-A50A-08BDA31F3257}
// *********************************************************************//
  IPCsExternalObjectDisp = dispinterface
    ['{E51DE5D6-F827-476C-A50A-08BDA31F3257}']
    property IsPicture: WordBool readonly dispid 301;
    property IsOleObject: WordBool readonly dispid 302;
    property Width: Integer dispid 303;
    property Height: Integer dispid 304;
    procedure SaveToFile(const Filename: WideString; Options: Integer); dispid 305;
    procedure SaveToStream(const Stream: IUnknown; Options: Integer); dispid 306;
    procedure LoadFromStream(const Stream: IUnknown); dispid 307;
    property Filename: WideString dispid 308;
    function CustomCommand(const Command: WideString; Params: OleVariant; out Results: OleVariant): WordBool; dispid 309;
    function IsErased: WordBool; dispid 310;
    procedure AssignTo(const Obj: IPCsBase); dispid 311;
    function Equals(const Obj: IPCsBase): WordBool; dispid 312;
    property IsValid: WordBool readonly dispid 313;
    property IsLinked: WordBool readonly dispid 314;
    property Visible: WordBool dispid 201;
    property Position: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant dispid 202;
    property LStatus: Integer dispid 203;
    function Move(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Figure: IPCsFigure): WordBool; dispid 204;
    procedure Rotate(DeltaAngle: Integer; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 205;
    procedure Scale(Factor: Single; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 208;
    property Handle: LongWord readonly dispid 206;
    procedure GetPosition(out X: Integer; out Y: Integer; out Z: Integer); dispid 207;
    procedure SetPosition(X: Integer; Y: Integer; Z: Integer); dispid 209;
    property State: Integer dispid 250;
    function ChangeDrawingOrder(AAction: psChangeDrawingOrder): WordBool; dispid 210;
  end;

// *********************************************************************//
// DispIntf:  IPCsExternalObjectEvents
// Flags:     (0)
// GUID:      {D703C1D4-B497-4CCE-8EF7-24BD10DDBAAF}
// *********************************************************************//
  IPCsExternalObjectEvents = dispinterface
    ['{D703C1D4-B497-4CCE-8EF7-24BD10DDBAAF}']
  end;

// *********************************************************************//
// Interface: IPCsExternalObject2
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {71A0B111-181F-4F28-9704-DFFF161FA5D8}
// *********************************************************************//
  IPCsExternalObject2 = interface(IPCsExternalObject)
    ['{71A0B111-181F-4F28-9704-DFFF161FA5D8}']
    procedure Delete; safecall;
    function Get_Name: WideString; safecall;
    procedure Set_Name(const Value: WideString); safecall;
    property Name: WideString read Get_Name write Set_Name;
  end;

// *********************************************************************//
// DispIntf:  IPCsExternalObject2Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {71A0B111-181F-4F28-9704-DFFF161FA5D8}
// *********************************************************************//
  IPCsExternalObject2Disp = dispinterface
    ['{71A0B111-181F-4F28-9704-DFFF161FA5D8}']
    procedure Delete; dispid 401;
    property Name: WideString dispid 402;
    property IsPicture: WordBool readonly dispid 301;
    property IsOleObject: WordBool readonly dispid 302;
    property Width: Integer dispid 303;
    property Height: Integer dispid 304;
    procedure SaveToFile(const Filename: WideString; Options: Integer); dispid 305;
    procedure SaveToStream(const Stream: IUnknown; Options: Integer); dispid 306;
    procedure LoadFromStream(const Stream: IUnknown); dispid 307;
    property Filename: WideString dispid 308;
    function CustomCommand(const Command: WideString; Params: OleVariant; out Results: OleVariant): WordBool; dispid 309;
    function IsErased: WordBool; dispid 310;
    procedure AssignTo(const Obj: IPCsBase); dispid 311;
    function Equals(const Obj: IPCsBase): WordBool; dispid 312;
    property IsValid: WordBool readonly dispid 313;
    property IsLinked: WordBool readonly dispid 314;
    property Visible: WordBool dispid 201;
    property Position: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant dispid 202;
    property LStatus: Integer dispid 203;
    function Move(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Figure: IPCsFigure): WordBool; dispid 204;
    procedure Rotate(DeltaAngle: Integer; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 205;
    procedure Scale(Factor: Single; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 208;
    property Handle: LongWord readonly dispid 206;
    procedure GetPosition(out X: Integer; out Y: Integer; out Z: Integer); dispid 207;
    procedure SetPosition(X: Integer; Y: Integer; Z: Integer); dispid 209;
    property State: Integer dispid 250;
    function ChangeDrawingOrder(AAction: psChangeDrawingOrder): WordBool; dispid 210;
  end;

// *********************************************************************//
// Interface: IPCsPicture
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {8985BD82-9EDB-4674-84BB-32A2D457299D}
// *********************************************************************//
  IPCsPicture = interface(IPCsExternalObject)
    ['{8985BD82-9EDB-4674-84BB-32A2D457299D}']
    function Get_IsEnhancedMetafile: WordBool; safecall;
    function Get_IsMetafile: WordBool; safecall;
    function Get_PictureHandle: Integer; safecall;
    property IsEnhancedMetafile: WordBool read Get_IsEnhancedMetafile;
    property IsMetafile: WordBool read Get_IsMetafile;
    property PictureHandle: Integer read Get_PictureHandle;
  end;

// *********************************************************************//
// DispIntf:  IPCsPictureDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {8985BD82-9EDB-4674-84BB-32A2D457299D}
// *********************************************************************//
  IPCsPictureDisp = dispinterface
    ['{8985BD82-9EDB-4674-84BB-32A2D457299D}']
    property IsEnhancedMetafile: WordBool readonly dispid 401;
    property IsMetafile: WordBool readonly dispid 402;
    property PictureHandle: Integer readonly dispid 403;
    property IsPicture: WordBool readonly dispid 301;
    property IsOleObject: WordBool readonly dispid 302;
    property Width: Integer dispid 303;
    property Height: Integer dispid 304;
    procedure SaveToFile(const Filename: WideString; Options: Integer); dispid 305;
    procedure SaveToStream(const Stream: IUnknown; Options: Integer); dispid 306;
    procedure LoadFromStream(const Stream: IUnknown); dispid 307;
    property Filename: WideString dispid 308;
    function CustomCommand(const Command: WideString; Params: OleVariant; out Results: OleVariant): WordBool; dispid 309;
    function IsErased: WordBool; dispid 310;
    procedure AssignTo(const Obj: IPCsBase); dispid 311;
    function Equals(const Obj: IPCsBase): WordBool; dispid 312;
    property IsValid: WordBool readonly dispid 313;
    property IsLinked: WordBool readonly dispid 314;
    property Visible: WordBool dispid 201;
    property Position: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant dispid 202;
    property LStatus: Integer dispid 203;
    function Move(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Figure: IPCsFigure): WordBool; dispid 204;
    procedure Rotate(DeltaAngle: Integer; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 205;
    procedure Scale(Factor: Single; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 208;
    property Handle: LongWord readonly dispid 206;
    procedure GetPosition(out X: Integer; out Y: Integer; out Z: Integer); dispid 207;
    procedure SetPosition(X: Integer; Y: Integer; Z: Integer); dispid 209;
    property State: Integer dispid 250;
    function ChangeDrawingOrder(AAction: psChangeDrawingOrder): WordBool; dispid 210;
  end;

// *********************************************************************//
// DispIntf:  IPCsPictureEvents
// Flags:     (0)
// GUID:      {6D87C31D-591E-4332-87B7-C4E49011266D}
// *********************************************************************//
  IPCsPictureEvents = dispinterface
    ['{6D87C31D-591E-4332-87B7-C4E49011266D}']
  end;

// *********************************************************************//
// Interface: IPCsDatabaseSetup
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {A843D727-CABB-4842-90B6-FA063DEDD883}
// *********************************************************************//
  IPCsDatabaseSetup = interface(IDispatch)
    ['{A843D727-CABB-4842-90B6-FA063DEDD883}']
    function Get_AssignedFieldNames(FieldType: psSetupDatabaseFieldTypes): WideString; safecall;
    procedure Set_AssignedFieldNames(FieldType: psSetupDatabaseFieldTypes; const Value: WideString); safecall;
    property AssignedFieldNames[FieldType: psSetupDatabaseFieldTypes]: WideString read Get_AssignedFieldNames write Set_AssignedFieldNames;
  end;

// *********************************************************************//
// DispIntf:  IPCsDatabaseSetupDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {A843D727-CABB-4842-90B6-FA063DEDD883}
// *********************************************************************//
  IPCsDatabaseSetupDisp = dispinterface
    ['{A843D727-CABB-4842-90B6-FA063DEDD883}']
    property AssignedFieldNames[FieldType: psSetupDatabaseFieldTypes]: WideString dispid 201;
  end;

// *********************************************************************//
// Interface: IPCsLine4
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {2865F44C-BA03-4D85-A263-8FF3C7B6D841}
// *********************************************************************//
  IPCsLine4 = interface(IPCsLine3)
    ['{2865F44C-BA03-4D85-A263-8FF3C7B6D841}']
    function Get_IsRoundLine: WordBool; safecall;
    procedure Set_IsRoundLine(Value: WordBool); safecall;
    function Get_LayerIndex: Integer; safecall;
    procedure Set_LayerIndex(Value: Integer); safecall;
    procedure SetQuantity(Quantity: Double); safecall;
    property IsRoundLine: WordBool read Get_IsRoundLine write Set_IsRoundLine;
    property LayerIndex: Integer read Get_LayerIndex write Set_LayerIndex;
  end;

// *********************************************************************//
// DispIntf:  IPCsLine4Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {2865F44C-BA03-4D85-A263-8FF3C7B6D841}
// *********************************************************************//
  IPCsLine4Disp = dispinterface
    ['{2865F44C-BA03-4D85-A263-8FF3C7B6D841}']
    property IsRoundLine: WordBool dispid 501;
    property LayerIndex: Integer dispid 505;
    procedure SetQuantity(Quantity: Double); dispid 502;
    procedure UniteJoinedLines; dispid 403;
    procedure RemoveExtendedSegments; dispid 404;
    procedure RemoveZeroVectors; dispid 405;
    function CanUniteWith(const ALine: IPCsLine): WordBool; dispid 406;
    function UniteWith(const ALine: IPCsLine): WordBool; dispid 407;
    function MoveLinePoint(PointIndex: Integer; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant;
                           Options: Integer): WordBool; dispid 314;
    function MoveLinePointXYZ(PointIndex: Integer; X: Integer; Y: Integer; Z: Integer;
                              Options: Integer): WordBool; dispid 320;
    property Application: IPCsApplication readonly dispid 321;
    function Equals(const Obj: IPCsBase): WordBool; dispid 322;
    function IsErased: WordBool; dispid 323;
    property Points: PCsPoints readonly dispid 401;
    procedure AddPoint(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 402;
    property IsElectrical: WordBool dispid 301;
    property LineStyle: PsLineStyle dispid 302;
    property Color: PsPenColor dispid 303;
    property PenWidth: Integer dispid 304;
    property Quantity: Double readonly dispid 305;
    property IsFilled: WordBool dispid 306;
    property IsJumperingLink: WordBool dispid 308;
    property IsCableWire: WordBool dispid 307;
    property IsPartLine: WordBool readonly dispid 310;
    property Extended: WordBool dispid 311;
    property IsOpen: WordBool dispid 312;
    property IsSignalsJoined: WordBool readonly dispid 313;
    function AddBendingPoint(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; BendFactor: Double): WordBool; dispid 315;
    procedure ConnectTo(const JoinPoint: IPCsJoinPoint); dispid 316;
    property Pattern: Integer dispid 317;
    property MultilineDistance: Integer dispid 318;
    property Layer: IPCsLayer dispid 319;
    property IsCurvedLine: WordBool dispid 309;
    property Visible: WordBool dispid 201;
    procedure AddPointXYZ(X: Integer; Y: Integer; Z: Integer); dispid 202;
    procedure Delete; dispid 203;
    property LineTexts: IPCsLineTexts readonly dispid 204;
    property Handle: Integer readonly dispid 205;
    property State: Integer dispid 250;
    property ObjectType: PsPCsObjectType readonly dispid 450;
    procedure InitLineTexts; dispid 206;
    function MoveRelative(AOffset: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Figure: IPCsFigure): WordBool; dispid 207;
    function MoveRelativeXYZ(OffsetX: Integer; OffsetY: Integer; OffsetZ: Integer;
                             const Figure: IPCsFigure): WordBool; dispid 208;
    function ChangeDrawingOrder(AAction: psChangeDrawingOrder): WordBool; dispid 209;
  end;

// *********************************************************************//
// Interface: IPCSLine5
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {DD427251-8AB8-4A46-A137-9BBBE3316984}
// *********************************************************************//
  IPCSLine5 = interface(IPCsLine4)
    ['{DD427251-8AB8-4A46-A137-9BBBE3316984}']
    function Get_RGBColor: TRGBColor; safecall;
    procedure Set_RGBColor(Value: TRGBColor); safecall;
    function Get_IsWithoutWirenumber: WordBool; safecall;
    procedure Set_IsWithoutWirenumber(Value: WordBool); safecall;
    property RGBColor: TRGBColor read Get_RGBColor write Set_RGBColor;
    property IsWithoutWirenumber: WordBool read Get_IsWithoutWirenumber write Set_IsWithoutWirenumber;
  end;

// *********************************************************************//
// DispIntf:  IPCSLine5Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {DD427251-8AB8-4A46-A137-9BBBE3316984}
// *********************************************************************//
  IPCSLine5Disp = dispinterface
    ['{DD427251-8AB8-4A46-A137-9BBBE3316984}']
    property RGBColor: TRGBColor dispid 601;
    property IsWithoutWirenumber: WordBool dispid 602;
    property IsRoundLine: WordBool dispid 501;
    property LayerIndex: Integer dispid 505;
    procedure SetQuantity(Quantity: Double); dispid 502;
    procedure UniteJoinedLines; dispid 403;
    procedure RemoveExtendedSegments; dispid 404;
    procedure RemoveZeroVectors; dispid 405;
    function CanUniteWith(const ALine: IPCsLine): WordBool; dispid 406;
    function UniteWith(const ALine: IPCsLine): WordBool; dispid 407;
    function MoveLinePoint(PointIndex: Integer; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant;
                           Options: Integer): WordBool; dispid 314;
    function MoveLinePointXYZ(PointIndex: Integer; X: Integer; Y: Integer; Z: Integer;
                              Options: Integer): WordBool; dispid 320;
    property Application: IPCsApplication readonly dispid 321;
    function Equals(const Obj: IPCsBase): WordBool; dispid 322;
    function IsErased: WordBool; dispid 323;
    property Points: PCsPoints readonly dispid 401;
    procedure AddPoint(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 402;
    property IsElectrical: WordBool dispid 301;
    property LineStyle: PsLineStyle dispid 302;
    property Color: PsPenColor dispid 303;
    property PenWidth: Integer dispid 304;
    property Quantity: Double readonly dispid 305;
    property IsFilled: WordBool dispid 306;
    property IsJumperingLink: WordBool dispid 308;
    property IsCableWire: WordBool dispid 307;
    property IsPartLine: WordBool readonly dispid 310;
    property Extended: WordBool dispid 311;
    property IsOpen: WordBool dispid 312;
    property IsSignalsJoined: WordBool readonly dispid 313;
    function AddBendingPoint(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; BendFactor: Double): WordBool; dispid 315;
    procedure ConnectTo(const JoinPoint: IPCsJoinPoint); dispid 316;
    property Pattern: Integer dispid 317;
    property MultilineDistance: Integer dispid 318;
    property Layer: IPCsLayer dispid 319;
    property IsCurvedLine: WordBool dispid 309;
    property Visible: WordBool dispid 201;
    procedure AddPointXYZ(X: Integer; Y: Integer; Z: Integer); dispid 202;
    procedure Delete; dispid 203;
    property LineTexts: IPCsLineTexts readonly dispid 204;
    property Handle: Integer readonly dispid 205;
    property State: Integer dispid 250;
    property ObjectType: PsPCsObjectType readonly dispid 450;
    procedure InitLineTexts; dispid 206;
    function MoveRelative(AOffset: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Figure: IPCsFigure): WordBool; dispid 207;
    function MoveRelativeXYZ(OffsetX: Integer; OffsetY: Integer; OffsetZ: Integer;
                             const Figure: IPCsFigure): WordBool; dispid 208;
    function ChangeDrawingOrder(AAction: psChangeDrawingOrder): WordBool; dispid 209;
  end;

// *********************************************************************//
// Interface: IPCsSymbol4
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {5E752712-E638-42D9-86EE-6002653969A7}
// *********************************************************************//
  IPCsSymbol4 = interface(IPCsSymbol3)
    ['{5E752712-E638-42D9-86EE-6002653969A7}']
    function Get_Component: IPCsComponent; safecall;
    function Get_Name: IPCsComponentName; safecall;
    procedure Set_Name(const Value: IPCsComponentName); safecall;
    procedure InitObjectRefID; safecall;
    property Component: IPCsComponent read Get_Component;
    property Name: IPCsComponentName read Get_Name write Set_Name;
  end;

// *********************************************************************//
// DispIntf:  IPCsSymbol4Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {5E752712-E638-42D9-86EE-6002653969A7}
// *********************************************************************//
  IPCsSymbol4Disp = dispinterface
    ['{5E752712-E638-42D9-86EE-6002653969A7}']
    property Component: IPCsComponent readonly dispid 601;
    property Name: IPCsComponentName dispid 602;
    procedure InitObjectRefID; dispid 603;
    property IsFeedThroughTerminal: WordBool readonly dispid 501;
    property IsTerminal: WordBool dispid 502;
    function IsSameComponent(const ASymbol: IPCsSymbol; Options: psIsSameComponentOptions): WordBool; dispid 503;
    function IsIdenticalTo(const ASymbol: IPCsSymbol): WordBool; dispid 504;
    property LayerIndex: Integer dispid 505;
    property IsElectrical: WordBool dispid 506;
    property IsMechanical: WordBool dispid 507;
    property IncludeInMechanicalLoad: WordBool dispid 508;
    property IsProtected: WordBool dispid 509;
    function IsSamePin(ConnectionIndex: Integer; const OtherSymbol: IPCsSymbol;
                       OtherConnectionIndex: Integer; Options: psIsSameComponentOptions): WordBool; dispid 510;
    function IsConnectionPair(ConnectionIndex1: Integer; ConnectionIndex2: Integer): WordBool; dispid 511;
    function HasSameNameAs(const ASymbol: IPCsSymbol): WordBool; dispid 512;
    function IsMetaconnected(ConnectionIndex: Integer; const OtherSymbol: IPCsSymbol;
                             OtherConnectionIndex: Integer; Options: Integer): WordBool; dispid 513;
    procedure SetQuantity(Quantity: Double); dispid 514;
    property SymbolData[Index: Integer]: WideString dispid 451;
    property SymbolDataCount: Integer readonly dispid 452;
    function Equals(const Obj: IPCsBase): WordBool; dispid 401;
    function IsErased: WordBool; dispid 403;
    property LibType: IPCsLibType readonly dispid 402;
    property IsVMirrored: WordBool dispid 405;
    property IsHMirrored: WordBool dispid 406;
    property SymbolTexts: IPCsSymbolTexts readonly dispid 407;
    property Page: IPCsPage readonly dispid 408;
    property Layer: IPCsLayer dispid 409;
    property FullName: WideString dispid 302;
    function GetRect(AType: PsSymbolRectType): {NOT_OLEAUTO(TPCsRect)}OleVariant; dispid 303;
    property Connections: IPCsConnections readonly dispid 304;
    property Rotation: Integer dispid 305;
    property ScaleFactor: Double dispid 306;
    function Explode(const ADestination: IPCsFigure): WordBool; dispid 307;
    property SymbolType1: psSymbolType dispid 301;
    property SymbolType2: psSymbolType dispid 308;
    procedure ApplicationDataWriteString(const ASection: WideString; const AIdent: WideString;
                                         const AValue: WideString); dispid 309;
    function ApplicationDataReadString(const ASection: WideString; const AIdent: WideString;
                                       const ADefault: WideString): WideString; dispid 310;
    property Datafields: IPCsDatafields readonly dispid 311;
    procedure Delete; dispid 312;
    property ObjectType: PsPCsObjectType readonly dispid 450;
    property ComponentGroupNo: Integer dispid 313;
    property ComponentPosNo: Integer dispid 314;
    property Reference: IPCsReference readonly dispid 315;
    property WithReference: WordBool dispid 316;
    function UpdateFromDatabase(Reserved: Integer): WordBool; dispid 317;
    property Visible: WordBool dispid 201;
    property Position: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant dispid 202;
    property LStatus: Integer dispid 203;
    function Move(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Figure: IPCsFigure): WordBool; dispid 204;
    procedure Rotate(DeltaAngle: Integer; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 205;
    procedure Scale(Factor: Single; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 208;
    property Handle: LongWord readonly dispid 206;
    procedure GetPosition(out X: Integer; out Y: Integer; out Z: Integer); dispid 207;
    procedure SetPosition(X: Integer; Y: Integer; Z: Integer); dispid 209;
    property State: Integer dispid 250;
    function ChangeDrawingOrder(AAction: psChangeDrawingOrder): WordBool; dispid 210;
  end;

// *********************************************************************//
// Interface: IPCsPage4
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {8CFB2E38-7691-4120-93AF-F0D48EAB9EFE}
// *********************************************************************//
  IPCsPage4 = interface(IPCsPage3)
    ['{8CFB2E38-7691-4120-93AF-F0D48EAB9EFE}']
    function Get_ReferenceDesignation: IPCsObjectRefID; safecall;
    procedure UpdateReference; safecall;
    function Get_RefDesString(Aspect: psReferenceDesignationAspect): WideString; safecall;
    procedure Set_RefDesString(Aspect: psReferenceDesignationAspect; const Value: WideString); safecall;
    property ReferenceDesignation: IPCsObjectRefID read Get_ReferenceDesignation;
    property RefDesString[Aspect: psReferenceDesignationAspect]: WideString read Get_RefDesString write Set_RefDesString;
  end;

// *********************************************************************//
// DispIntf:  IPCsPage4Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {8CFB2E38-7691-4120-93AF-F0D48EAB9EFE}
// *********************************************************************//
  IPCsPage4Disp = dispinterface
    ['{8CFB2E38-7691-4120-93AF-F0D48EAB9EFE}']
    property ReferenceDesignation: IPCsObjectRefID readonly dispid 501;
    procedure UpdateReference; dispid 301;
    property RefDesString[Aspect: psReferenceDesignationAspect]: WideString dispid 302;
    procedure GetDrawingArea(out X1: Integer; out Y1: Integer; out X2: Integer; out Y2: Integer); dispid 502;
    property GridSize: Integer dispid 504;
    property HighlightedObjectsCount[Kind: psHighlighting; Options: psIterateDrawingObjectsOptions]: Integer readonly dispid 202;
    property LastModifiedTime: Double dispid 201;
    function PlaceSymbolXYZ(X: Integer; Y: Integer; Z: Integer; const Filename: WideString): IPCsSymbol3; dispid 203;
  end;

// *********************************************************************//
// Interface: IPCsText2
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {36BAF7D3-3BB8-4F52-9E85-6CD291ED2CA9}
// *********************************************************************//
  IPCsText2 = interface(IPCsText)
    ['{36BAF7D3-3BB8-4F52-9E85-6CD291ED2CA9}']
    function Get_RGBColor: Integer; safecall;
    procedure Set_RGBColor(Value: Integer); safecall;
    property RGBColor: Integer read Get_RGBColor write Set_RGBColor;
  end;

// *********************************************************************//
// DispIntf:  IPCsText2Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {36BAF7D3-3BB8-4F52-9E85-6CD291ED2CA9}
// *********************************************************************//
  IPCsText2Disp = dispinterface
    ['{36BAF7D3-3BB8-4F52-9E85-6CD291ED2CA9}']
    property RGBColor: Integer dispid 409;
    property Value: WideString dispid 401;
    property Origin: PsTextOrigin dispid 402;
    property Font: IPCsTextFont readonly dispid 403;
    property TextType: PsTextOrigin readonly dispid 404;
    property FieldWidth: Integer dispid 405;
    property WrapText: WordBool dispid 406;
    property LinkID: Integer dispid 407;
    property PositionExt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant dispid 408;
    property Color: PsPenColor dispid 301;
    property Layer: IPCsLayer dispid 302;
    property DisplayText: WideString readonly dispid 303;
    property LineCount: Integer readonly dispid 304;
    property Rotation: Integer dispid 305;
    property IsBlank: WordBool readonly dispid 306;
    function GetRect(AFlags: psTextRectFlags): {NOT_OLEAUTO(TPCsRect)}OleVariant; dispid 307;
    procedure Delete; dispid 308;
    property ObjectType: PsPCsObjectType readonly dispid 450;
    property NotTranslated: WordBool dispid 309;
    property Visible: WordBool dispid 201;
    property Position: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant dispid 202;
    property LStatus: Integer dispid 203;
    function Move(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Figure: IPCsFigure): WordBool; dispid 204;
    procedure Rotate(DeltaAngle: Integer; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 205;
    procedure Scale(Factor: Single; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 208;
    property Handle: LongWord readonly dispid 206;
    procedure GetPosition(out X: Integer; out Y: Integer; out Z: Integer); dispid 207;
    procedure SetPosition(X: Integer; Y: Integer; Z: Integer); dispid 209;
    property State: Integer dispid 250;
    function ChangeDrawingOrder(AAction: psChangeDrawingOrder): WordBool; dispid 210;
  end;

// *********************************************************************//
// Interface: IPCsNameText4
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {59F989D2-211C-48DF-9B25-D445CD7C1C01}
// *********************************************************************//
  IPCsNameText4 = interface(IPCsNameText3)
    ['{59F989D2-211C-48DF-9B25-D445CD7C1C01}']
    function Get_AsString(const Format: IPCsRefIDFormat): WideString; safecall;
    procedure Set_AsString(const Format: IPCsRefIDFormat; const Value: WideString); safecall;
    property AsString[const Format: IPCsRefIDFormat]: WideString read Get_AsString write Set_AsString;
  end;

// *********************************************************************//
// DispIntf:  IPCsNameText4Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {59F989D2-211C-48DF-9B25-D445CD7C1C01}
// *********************************************************************//
  IPCsNameText4Disp = dispinterface
    ['{59F989D2-211C-48DF-9B25-D445CD7C1C01}']
    property AsString[const Format: IPCsRefIDFormat]: WideString dispid 701;
    property RefDesString[Aspect: psReferenceDesignationAspect]: WideString dispid 600;
    property ReferenceDesignation: IPCsObjectRefID readonly dispid 601;
    property DisplayText2: WideString readonly dispid 501;
    function Equals(const Obj: IPCsBase): WordBool; dispid 502;
    property Application: IPCsApplication2 readonly dispid 503;
    function IsErased: WordBool; dispid 504;
    property RefDesFunction: WideString dispid 409;
    property RefDesLocation: WideString dispid 410;
    property RefDesOptions: psSymbolRefDesOption dispid 411;
    property Value: WideString dispid 401;
    property Origin: PsTextOrigin dispid 402;
    property Font: IPCsTextFont readonly dispid 403;
    property TextType: PsTextOrigin readonly dispid 404;
    property FieldWidth: Integer dispid 405;
    property WrapText: WordBool dispid 406;
    property LinkID: Integer dispid 407;
    property PositionExt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant dispid 408;
    property Color: PsPenColor dispid 301;
    property Layer: IPCsLayer dispid 302;
    property DisplayText: WideString readonly dispid 303;
    property LineCount: Integer readonly dispid 304;
    property Rotation: Integer dispid 305;
    property IsBlank: WordBool readonly dispid 306;
    function GetRect(AFlags: psTextRectFlags): {NOT_OLEAUTO(TPCsRect)}OleVariant; dispid 307;
    procedure Delete; dispid 308;
    property ObjectType: PsPCsObjectType readonly dispid 450;
    property NotTranslated: WordBool dispid 309;
    property Visible: WordBool dispid 201;
    property Position: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant dispid 202;
    property LStatus: Integer dispid 203;
    function Move(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Figure: IPCsFigure): WordBool; dispid 204;
    procedure Rotate(DeltaAngle: Integer; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 205;
    procedure Scale(Factor: Single; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 208;
    property Handle: LongWord readonly dispid 206;
    procedure GetPosition(out X: Integer; out Y: Integer; out Z: Integer); dispid 207;
    procedure SetPosition(X: Integer; Y: Integer; Z: Integer); dispid 209;
    property State: Integer dispid 250;
    function ChangeDrawingOrder(AAction: psChangeDrawingOrder): WordBool; dispid 210;
  end;

// *********************************************************************//
// Interface: IPCsComponent
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {8DC3B9F5-5E29-48CE-8DBE-C700B1280769}
// *********************************************************************//
  IPCsComponent = interface(IDispatch)
    ['{8DC3B9F5-5E29-48CE-8DBE-C700B1280769}']
    function Get_Symbols: IPCsComponentSymbols; safecall;
    function Get_Name: IPCsComponentName; safecall;
    property Symbols: IPCsComponentSymbols read Get_Symbols;
    property Name: IPCsComponentName read Get_Name;
  end;

// *********************************************************************//
// DispIntf:  IPCsComponentDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {8DC3B9F5-5E29-48CE-8DBE-C700B1280769}
// *********************************************************************//
  IPCsComponentDisp = dispinterface
    ['{8DC3B9F5-5E29-48CE-8DBE-C700B1280769}']
    property Symbols: IPCsComponentSymbols readonly dispid 201;
    property Name: IPCsComponentName readonly dispid 202;
  end;

// *********************************************************************//
// DispIntf:  IPCsComponentEvents
// Flags:     (0)
// GUID:      {6DB9B522-43A8-499F-A47E-6CB900911C56}
// *********************************************************************//
  IPCsComponentEvents = dispinterface
    ['{6DB9B522-43A8-499F-A47E-6CB900911C56}']
  end;

// *********************************************************************//
// Interface: IPCsComponent2
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {9D293E30-A1E2-47AB-9E73-DD99425AE035}
// *********************************************************************//
  IPCsComponent2 = interface(IPCsComponent)
    ['{9D293E30-A1E2-47AB-9E73-DD99425AE035}']
    procedure ChangeName(const AName: IPCsCompName); safecall;
    function Get_GroupNo: Integer; safecall;
    procedure Set_GroupNo(Value: Integer); safecall;
    property GroupNo: Integer read Get_GroupNo write Set_GroupNo;
  end;

// *********************************************************************//
// DispIntf:  IPCsComponent2Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {9D293E30-A1E2-47AB-9E73-DD99425AE035}
// *********************************************************************//
  IPCsComponent2Disp = dispinterface
    ['{9D293E30-A1E2-47AB-9E73-DD99425AE035}']
    procedure ChangeName(const AName: IPCsCompName); dispid 301;
    property GroupNo: Integer dispid 302;
    property Symbols: IPCsComponentSymbols readonly dispid 201;
    property Name: IPCsComponentName readonly dispid 202;
  end;

// *********************************************************************//
// Interface: IPCsComponentName
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {7CD5EEA6-542C-48AF-894F-7D743883E37C}
// *********************************************************************//
  IPCsComponentName = interface(IPCsBase)
    ['{7CD5EEA6-542C-48AF-894F-7D743883E37C}']
    function Get_ReferenceDesignation: IPCsObjectRefID; safecall;
    function Get_ComponentName: WideString; safecall;
    procedure Set_ComponentName(const Value: WideString); safecall;
    function Get_FullName: WideString; safecall;
    procedure Set_FullName(const Value: WideString); safecall;
    function Get_FullName2: WideString; safecall;
    property ReferenceDesignation: IPCsObjectRefID read Get_ReferenceDesignation;
    property ComponentName: WideString read Get_ComponentName write Set_ComponentName;
    property FullName: WideString read Get_FullName write Set_FullName;
    property FullName2: WideString read Get_FullName2;
  end;

// *********************************************************************//
// DispIntf:  IPCsComponentNameDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {7CD5EEA6-542C-48AF-894F-7D743883E37C}
// *********************************************************************//
  IPCsComponentNameDisp = dispinterface
    ['{7CD5EEA6-542C-48AF-894F-7D743883E37C}']
    property ReferenceDesignation: IPCsObjectRefID readonly dispid 401;
    property ComponentName: WideString dispid 402;
    property FullName: WideString dispid 301;
    property FullName2: WideString readonly dispid 302;
    property Handle: LongWord readonly dispid 201;
    function CustomCommand(const Command: WideString; Params: OleVariant; out Results: OleVariant): WordBool; dispid 202;
    function Equals(const Obj: IPCsBase): WordBool; dispid 203;
    procedure AssignTo(const Obj: IPCsBase); dispid 204;
    property IsErased: WordBool readonly dispid 205;
    property Application: IPCsApplication2 readonly dispid 206;
    property LStatus: Integer readonly dispid 207;
  end;

// *********************************************************************//
// Interface: IPCsObjectRefID
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {92E3FB02-7A93-4CEB-BB4B-E74E2DD7663A}
// *********************************************************************//
  IPCsObjectRefID = interface(IPCsBase)
    ['{92E3FB02-7A93-4CEB-BB4B-E74E2DD7663A}']
    function Get_Aspects(Aspect: psReferenceDesignationAspect): IPCsReferenceDesignation; safecall;
    procedure Set_Aspects(Aspect: psReferenceDesignationAspect;
                          const Value: IPCsReferenceDesignation); safecall;
    property Aspects[Aspect: psReferenceDesignationAspect]: IPCsReferenceDesignation read Get_Aspects write Set_Aspects;
  end;

// *********************************************************************//
// DispIntf:  IPCsObjectRefIDDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {92E3FB02-7A93-4CEB-BB4B-E74E2DD7663A}
// *********************************************************************//
  IPCsObjectRefIDDisp = dispinterface
    ['{92E3FB02-7A93-4CEB-BB4B-E74E2DD7663A}']
    property Aspects[Aspect: psReferenceDesignationAspect]: IPCsReferenceDesignation dispid 102;
    property Handle: LongWord readonly dispid 201;
    function CustomCommand(const Command: WideString; Params: OleVariant; out Results: OleVariant): WordBool; dispid 202;
    function Equals(const Obj: IPCsBase): WordBool; dispid 203;
    procedure AssignTo(const Obj: IPCsBase); dispid 204;
    property IsErased: WordBool readonly dispid 205;
    property Application: IPCsApplication2 readonly dispid 206;
    property LStatus: Integer readonly dispid 207;
  end;

// *********************************************************************//
// Interface: IPCsReferenceDesignations3
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {8500668B-9672-416D-93C1-E55856AB7CAA}
// *********************************************************************//
  IPCsReferenceDesignations3 = interface(IPCsReferenceDesignations2)
    ['{8500668B-9672-416D-93C1-E55856AB7CAA}']
    function Get_DisplayFormat: IPCsRefIDFormat; safecall;
    function Get_ImportFormat: IPCsRefIDFormat; safecall;
    property DisplayFormat: IPCsRefIDFormat read Get_DisplayFormat;
    property ImportFormat: IPCsRefIDFormat read Get_ImportFormat;
  end;

// *********************************************************************//
// DispIntf:  IPCsReferenceDesignations3Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {8500668B-9672-416D-93C1-E55856AB7CAA}
// *********************************************************************//
  IPCsReferenceDesignations3Disp = dispinterface
    ['{8500668B-9672-416D-93C1-E55856AB7CAA}']
    property DisplayFormat: IPCsRefIDFormat readonly dispid 402;
    property ImportFormat: IPCsRefIDFormat readonly dispid 403;
    property Aspects[Aspect: psReferenceDesignationAspect]: IPCsReferenceDesignationList readonly dispid 401;
    procedure Add(Aspect: psReferenceDesignationAspect; const Designation: WideString;
                  const Description: WideString); dispid 201;
    property Location[Index: Integer]: IPCsReferenceDesignation readonly dispid 202;
    property Func[Index: Integer]: IPCsReferenceDesignation readonly dispid 203;
    property LocationCount: Integer readonly dispid 204;
    property FunctionCount: Integer readonly dispid 205;
    function Find(Aspect: psReferenceDesignationAspect; const Designation: WideString): IPCsReferenceDesignation; dispid 206;
  end;

// *********************************************************************//
// Interface: IPCsReferenceDesignationList
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {89CE5B73-CFDF-40BD-BF28-9B00E5B56E48}
// *********************************************************************//
  IPCsReferenceDesignationList = interface(IDispatch)
    ['{89CE5B73-CFDF-40BD-BF28-9B00E5B56E48}']
    function Get_Count: Integer; safecall;
    function Get_Item(Index: Integer): IPCsReferenceDesignation; safecall;
    property Count: Integer read Get_Count;
    property Item[Index: Integer]: IPCsReferenceDesignation read Get_Item;
  end;

// *********************************************************************//
// DispIntf:  IPCsReferenceDesignationListDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {89CE5B73-CFDF-40BD-BF28-9B00E5B56E48}
// *********************************************************************//
  IPCsReferenceDesignationListDisp = dispinterface
    ['{89CE5B73-CFDF-40BD-BF28-9B00E5B56E48}']
    property Count: Integer readonly dispid 201;
    property Item[Index: Integer]: IPCsReferenceDesignation readonly dispid 202;
  end;

// *********************************************************************//
// Interface: IPCsLine6
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {99E2962E-0B32-44B3-B249-BFBAF6FC7822}
// *********************************************************************//
  IPCsLine6 = interface(IPCSLine5)
    ['{99E2962E-0B32-44B3-B249-BFBAF6FC7822}']
    function Get_Datafields: IPCsLineDatafields; safecall;
    property Datafields: IPCsLineDatafields read Get_Datafields;
  end;

// *********************************************************************//
// DispIntf:  IPCsLine6Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {99E2962E-0B32-44B3-B249-BFBAF6FC7822}
// *********************************************************************//
  IPCsLine6Disp = dispinterface
    ['{99E2962E-0B32-44B3-B249-BFBAF6FC7822}']
    property Datafields: IPCsLineDatafields readonly dispid 701;
    property RGBColor: TRGBColor dispid 601;
    property IsWithoutWirenumber: WordBool dispid 602;
    property IsRoundLine: WordBool dispid 501;
    property LayerIndex: Integer dispid 505;
    procedure SetQuantity(Quantity: Double); dispid 502;
    procedure UniteJoinedLines; dispid 403;
    procedure RemoveExtendedSegments; dispid 404;
    procedure RemoveZeroVectors; dispid 405;
    function CanUniteWith(const ALine: IPCsLine): WordBool; dispid 406;
    function UniteWith(const ALine: IPCsLine): WordBool; dispid 407;
    function MoveLinePoint(PointIndex: Integer; Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant;
                           Options: Integer): WordBool; dispid 314;
    function MoveLinePointXYZ(PointIndex: Integer; X: Integer; Y: Integer; Z: Integer;
                              Options: Integer): WordBool; dispid 320;
    property Application: IPCsApplication readonly dispid 321;
    function Equals(const Obj: IPCsBase): WordBool; dispid 322;
    function IsErased: WordBool; dispid 323;
    property Points: PCsPoints readonly dispid 401;
    procedure AddPoint(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant); dispid 402;
    property IsElectrical: WordBool dispid 301;
    property LineStyle: PsLineStyle dispid 302;
    property Color: PsPenColor dispid 303;
    property PenWidth: Integer dispid 304;
    property Quantity: Double readonly dispid 305;
    property IsFilled: WordBool dispid 306;
    property IsJumperingLink: WordBool dispid 308;
    property IsCableWire: WordBool dispid 307;
    property IsPartLine: WordBool readonly dispid 310;
    property Extended: WordBool dispid 311;
    property IsOpen: WordBool dispid 312;
    property IsSignalsJoined: WordBool readonly dispid 313;
    function AddBendingPoint(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; BendFactor: Double): WordBool; dispid 315;
    procedure ConnectTo(const JoinPoint: IPCsJoinPoint); dispid 316;
    property Pattern: Integer dispid 317;
    property MultilineDistance: Integer dispid 318;
    property Layer: IPCsLayer dispid 319;
    property IsCurvedLine: WordBool dispid 309;
    property Visible: WordBool dispid 201;
    procedure AddPointXYZ(X: Integer; Y: Integer; Z: Integer); dispid 202;
    procedure Delete; dispid 203;
    property LineTexts: IPCsLineTexts readonly dispid 204;
    property Handle: Integer readonly dispid 205;
    property State: Integer dispid 250;
    property ObjectType: PsPCsObjectType readonly dispid 450;
    procedure InitLineTexts; dispid 206;
    function MoveRelative(AOffset: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Figure: IPCsFigure): WordBool; dispid 207;
    function MoveRelativeXYZ(OffsetX: Integer; OffsetY: Integer; OffsetZ: Integer;
                             const Figure: IPCsFigure): WordBool; dispid 208;
    function ChangeDrawingOrder(AAction: psChangeDrawingOrder): WordBool; dispid 209;
  end;

// *********************************************************************//
// Interface: IPCsSymbolsEnumerator
// Flags:     (320) Dual OleAutomation
// GUID:      {A52328DD-0BBB-4923-AA9E-B7258F6FCB75}
// *********************************************************************//
  IPCsSymbolsEnumerator = interface(IEnumVARIANT)
    ['{A52328DD-0BBB-4923-AA9E-B7258F6FCB75}']
    function MoveNext: WordBool; safecall;
    function Get_Current: IPCsSymbol; safecall;
    property Current: IPCsSymbol read Get_Current;
  end;

// *********************************************************************//
// DispIntf:  IPCsSymbolsEnumeratorDisp
// Flags:     (320) Dual OleAutomation
// GUID:      {A52328DD-0BBB-4923-AA9E-B7258F6FCB75}
// *********************************************************************//
  IPCsSymbolsEnumeratorDisp = dispinterface
    ['{A52328DD-0BBB-4923-AA9E-B7258F6FCB75}']
    function MoveNext: WordBool; dispid 201;
    property Current: IPCsSymbol readonly dispid 202;
    function Next(celt: LongWord; const rgvar: OleVariant; out pceltFetched: LongWord): HResult; dispid 1610678272;
    function Skip(celt: LongWord): HResult; dispid 1610678273;
    function Reset: HResult; dispid 1610678274;
    function Clone(out ppenum: IEnumVARIANT): HResult; dispid 1610678275;
  end;

// *********************************************************************//
// Interface: IPCsComponentSymbols
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {C1E7B59A-00DF-45CF-93D2-A665A0F4A5A7}
// *********************************************************************//
  IPCsComponentSymbols = interface(IPCsBase)
    ['{C1E7B59A-00DF-45CF-93D2-A665A0F4A5A7}']
    function Get_Count: Integer; safecall;
    function Get_Item(Index: Integer): IPCsSymbol4; safecall;
    function GetEnumerator: IPCsSymbolsEnumerator; safecall;
    procedure Add(const Symbol: IPCsSymbol); safecall;
    property Count: Integer read Get_Count;
    property Item[Index: Integer]: IPCsSymbol4 read Get_Item;
  end;

// *********************************************************************//
// DispIntf:  IPCsComponentSymbolsDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {C1E7B59A-00DF-45CF-93D2-A665A0F4A5A7}
// *********************************************************************//
  IPCsComponentSymbolsDisp = dispinterface
    ['{C1E7B59A-00DF-45CF-93D2-A665A0F4A5A7}']
    property Count: Integer readonly dispid 401;
    property Item[Index: Integer]: IPCsSymbol4 readonly dispid 402;
    function GetEnumerator: IPCsSymbolsEnumerator; dispid 403;
    procedure Add(const Symbol: IPCsSymbol); dispid 404;
    property Handle: LongWord readonly dispid 201;
    function CustomCommand(const Command: WideString; Params: OleVariant; out Results: OleVariant): WordBool; dispid 202;
    function Equals(const Obj: IPCsBase): WordBool; dispid 203;
    procedure AssignTo(const Obj: IPCsBase); dispid 204;
    property IsErased: WordBool readonly dispid 205;
    property Application: IPCsApplication2 readonly dispid 206;
    property LStatus: Integer readonly dispid 207;
  end;

// *********************************************************************//
// Interface: IPCsPage5
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {2445EFE9-48A5-4CB2-A661-0C19EB145822}
// *********************************************************************//
  IPCsPage5 = interface(IPCsPage4)
    ['{2445EFE9-48A5-4CB2-A661-0C19EB145822}']
    function Get_Variables(Group: Integer; ID: Integer): WideString; safecall;
    function GetVariableIDs(Group: Integer): OleVariant; safecall;
    procedure SetVariable(Group: Integer; ID: Integer; const Value: WideString;
                          const RefIDFormat: IPCsRefIDFormat); safecall;
    procedure SetVariableList(Group: Integer; IDValuePairArray: OleVariant;
                              const RefIDFormat: IPCsRefIDFormat); safecall;
    function GetVariableList(Group: Integer): OleVariant; safecall;
    property Variables[Group: Integer; ID: Integer]: WideString read Get_Variables;
  end;

// *********************************************************************//
// DispIntf:  IPCsPage5Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {2445EFE9-48A5-4CB2-A661-0C19EB145822}
// *********************************************************************//
  IPCsPage5Disp = dispinterface
    ['{2445EFE9-48A5-4CB2-A661-0C19EB145822}']
    property Variables[Group: Integer; ID: Integer]: WideString readonly dispid 402;
    function GetVariableIDs(Group: Integer): OleVariant; dispid 403;
    procedure SetVariable(Group: Integer; ID: Integer; const Value: WideString;
                          const RefIDFormat: IPCsRefIDFormat); dispid 404;
    procedure SetVariableList(Group: Integer; IDValuePairArray: OleVariant;
                              const RefIDFormat: IPCsRefIDFormat); dispid 401;
    function GetVariableList(Group: Integer): OleVariant; dispid 405;
    property ReferenceDesignation: IPCsObjectRefID readonly dispid 501;
    procedure UpdateReference; dispid 301;
    property RefDesString[Aspect: psReferenceDesignationAspect]: WideString dispid 302;
    procedure GetDrawingArea(out X1: Integer; out Y1: Integer; out X2: Integer; out Y2: Integer); dispid 502;
    property GridSize: Integer dispid 504;
    property HighlightedObjectsCount[Kind: psHighlighting; Options: psIterateDrawingObjectsOptions]: Integer readonly dispid 202;
    property LastModifiedTime: Double dispid 201;
    function PlaceSymbolXYZ(X: Integer; Y: Integer; Z: Integer; const Filename: WideString): IPCsSymbol3; dispid 203;
  end;

// *********************************************************************//
// Interface: IPCsLineDatafields
// Flags:     (320) Dual OleAutomation
// GUID:      {4FD608C3-5B5C-42E1-B5E9-4021660091B1}
// *********************************************************************//
  IPCsLineDatafields = interface(IPCsDatafields)
    ['{4FD608C3-5B5C-42E1-B5E9-4021660091B1}']
  end;

// *********************************************************************//
// DispIntf:  IPCsLineDatafieldsDisp
// Flags:     (320) Dual OleAutomation
// GUID:      {4FD608C3-5B5C-42E1-B5E9-4021660091B1}
// *********************************************************************//
  IPCsLineDatafieldsDisp = dispinterface
    ['{4FD608C3-5B5C-42E1-B5E9-4021660091B1}']
    function Add(Pt: {NOT_OLEAUTO(TPCsPoint3D)}OleVariant; const Value: WideString): IPCsDatafield; dispid 101;
    property Item[Index: Integer]: IPCsDatafield readonly dispid 0; default;
    property Count: Integer readonly dispid 102;
    function FindByValue(const Value: WideString): IPCsDatafield; dispid 103;
    property _NewEnum: IEnumVARIANT readonly dispid -4;
  end;

// *********************************************************************//
// Interface: IPCsRefIDFormat
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {A40E9EA5-B27B-4FC8-9470-C036E0C1FDB0}
// *********************************************************************//
  IPCsRefIDFormat = interface(IDispatch)
    ['{A40E9EA5-B27B-4FC8-9470-C036E0C1FDB0}']
    function Get_AsString: WideString; safecall;
    procedure Set_AsString(const Value: WideString); safecall;
    property AsString: WideString read Get_AsString write Set_AsString;
  end;

// *********************************************************************//
// DispIntf:  IPCsRefIDFormatDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {A40E9EA5-B27B-4FC8-9470-C036E0C1FDB0}
// *********************************************************************//
  IPCsRefIDFormatDisp = dispinterface
    ['{A40E9EA5-B27B-4FC8-9470-C036E0C1FDB0}']
    property AsString: WideString dispid 201;
  end;

// *********************************************************************//
// Interface: IPCsCompName
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0C073709-BCC6-4134-A670-6B2F61879C3E}
// *********************************************************************//
  IPCsCompName = interface(IDispatch)
    ['{0C073709-BCC6-4134-A670-6B2F61879C3E}']
    function Get_ComponentName: WideString; safecall;
    procedure Set_ComponentName(const Value: WideString); safecall;
    function Get_ReferenceDesignation: IPCsObjectRefID; safecall;
    property ComponentName: WideString read Get_ComponentName write Set_ComponentName;
    property ReferenceDesignation: IPCsObjectRefID read Get_ReferenceDesignation;
  end;

// *********************************************************************//
// DispIntf:  IPCsCompNameDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0C073709-BCC6-4134-A670-6B2F61879C3E}
// *********************************************************************//
  IPCsCompNameDisp = dispinterface
    ['{0C073709-BCC6-4134-A670-6B2F61879C3E}']
    property ComponentName: WideString dispid 201;
    property ReferenceDesignation: IPCsObjectRefID readonly dispid 202;
  end;

// *********************************************************************//
// The Class CoPCsPoints provides a Create and CreateRemote method to
// create instances of the default interface IPCsPoints exposed by
// the CoClass PCsPoints. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsPoints = class
    class function Create: IPCsPoints;
    class function CreateRemote(const MachineName: string): IPCsPoints;
  end;

// *********************************************************************//
// The Class CoPCsDocuments provides a Create and CreateRemote method to
// create instances of the default interface IPCsDocuments exposed by
// the CoClass PCsDocuments. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsDocuments = class
    class function Create: IPCsDocuments;
    class function CreateRemote(const MachineName: string): IPCsDocuments;
  end;

// *********************************************************************//
// The Class CoPCsApplicationPreferences provides a Create and CreateRemote method to
// create instances of the default interface IPCsApplicationPreferences exposed by
// the CoClass PCsApplicationPreferences. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsApplicationPreferences = class
    class function Create: IPCsApplicationPreferences;
    class function CreateRemote(const MachineName: string): IPCsApplicationPreferences;
  end;

// *********************************************************************//
// The Class CoPCsStatusPoint provides a Create and CreateRemote method to
// create instances of the default interface IPCsStatusPoint exposed by
// the CoClass PCsStatusPoint. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsStatusPoint = class
    class function Create: IPCsStatusPoint;
    class function CreateRemote(const MachineName: string): IPCsStatusPoint;
  end;

// *********************************************************************//
// The Class CoPCsLibTypes provides a Create and CreateRemote method to
// create instances of the default interface IPCsLibTypes exposed by
// the CoClass PCsLibTypes. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsLibTypes = class
    class function Create: IPCsLibTypes;
    class function CreateRemote(const MachineName: string): IPCsLibTypes;
  end;

// *********************************************************************//
// The Class CoPCsLines provides a Create and CreateRemote method to
// create instances of the default interface IPCsLines exposed by
// the CoClass PCsLines. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsLines = class
    class function Create: IPCsLines;
    class function CreateRemote(const MachineName: string): IPCsLines;
  end;

// *********************************************************************//
// The Class CoPCsArcs provides a Create and CreateRemote method to
// create instances of the default interface IPCsArcs exposed by
// the CoClass PCsArcs. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsArcs = class
    class function Create: IPCsArcs;
    class function CreateRemote(const MachineName: string): IPCsArcs;
  end;

// *********************************************************************//
// The Class CoPCsPageSetup provides a Create and CreateRemote method to
// create instances of the default interface IPCsPageSetup exposed by
// the CoClass PCsPageSetup. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsPageSetup = class
    class function Create: IPCsPageSetup;
    class function CreateRemote(const MachineName: string): IPCsPageSetup;
  end;

// *********************************************************************//
// The Class CoPCsPageReference provides a Create and CreateRemote method to
// create instances of the default interface IPCsPageReference exposed by
// the CoClass PCsPageReference. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsPageReference = class
    class function Create: IPCsPageReference;
    class function CreateRemote(const MachineName: string): IPCsPageReference;
  end;

// *********************************************************************//
// The Class CoPCsBasePageRef provides a Create and CreateRemote method to
// create instances of the default interface IPCsBasePageRef exposed by
// the CoClass PCsBasePageRef. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsBasePageRef = class
    class function Create: IPCsBasePageRef;
    class function CreateRemote(const MachineName: string): IPCsBasePageRef;
  end;

// *********************************************************************//
// The Class CoPCsTexts provides a Create and CreateRemote method to
// create instances of the default interface IPCsTexts exposed by
// the CoClass PCsTexts. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsTexts = class
    class function Create: IPCsTexts;
    class function CreateRemote(const MachineName: string): IPCsTexts;
  end;

// *********************************************************************//
// The Class CoPCsSymbolTexts provides a Create and CreateRemote method to
// create instances of the default interface IPCsSymbolTexts exposed by
// the CoClass PCsSymbolTexts. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsSymbolTexts = class
    class function Create: IPCsSymbolTexts;
    class function CreateRemote(const MachineName: string): IPCsSymbolTexts;
  end;

// *********************************************************************//
// The Class CoPCsSymbolMenu provides a Create and CreateRemote method to
// create instances of the default interface IPCsSymbolMenu exposed by
// the CoClass PCsSymbolMenu. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsSymbolMenu = class
    class function Create: IPCsSymbolMenu;
    class function CreateRemote(const MachineName: string): IPCsSymbolMenu;
  end;

// *********************************************************************//
// The Class CoPCsLibTypeConnections provides a Create and CreateRemote method to
// create instances of the default interface IPCsLibTypeConnections exposed by
// the CoClass PCsLibTypeConnections. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsLibTypeConnections = class
    class function Create: IPCsLibTypeConnections;
    class function CreateRemote(const MachineName: string): IPCsLibTypeConnections;
  end;

// *********************************************************************//
// The Class CoPCsConnections provides a Create and CreateRemote method to
// create instances of the default interface IPCsConnections exposed by
// the CoClass PCsConnections. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsConnections = class
    class function Create: IPCsConnections;
    class function CreateRemote(const MachineName: string): IPCsConnections;
  end;

// *********************************************************************//
// The Class CoPCsConnectionTexts provides a Create and CreateRemote method to
// create instances of the default interface IPCsConnectionTexts exposed by
// the CoClass PCsConnectionTexts. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsConnectionTexts = class
    class function Create: IPCsConnectionTexts;
    class function CreateRemote(const MachineName: string): IPCsConnectionTexts;
  end;

// *********************************************************************//
// The Class CoPCsTextFont provides a Create and CreateRemote method to
// create instances of the default interface IPCsTextFont exposed by
// the CoClass PCsTextFont. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsTextFont = class
    class function Create: IPCsTextFont;
    class function CreateRemote(const MachineName: string): IPCsTextFont;
  end;

// *********************************************************************//
// The Class CoPCsJoinPoint provides a Create and CreateRemote method to
// create instances of the default interface IPCsJoinPoint exposed by
// the CoClass PCsJoinPoint. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsJoinPoint = class
    class function Create: IPCsJoinPoint;
    class function CreateRemote(const MachineName: string): IPCsJoinPoint;
  end;

// *********************************************************************//
// The Class CoPCsLayers provides a Create and CreateRemote method to
// create instances of the default interface IPCsLayers exposed by
// the CoClass PCsLayers. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsLayers = class
    class function Create: IPCsLayers;
    class function CreateRemote(const MachineName: string): IPCsLayers;
  end;

// *********************************************************************//
// The Class CoPCsLayer provides a Create and CreateRemote method to
// create instances of the default interface IPCsLayer exposed by
// the CoClass PCsLayer. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsLayer = class
    class function Create: IPCsLayer;
    class function CreateRemote(const MachineName: string): IPCsLayer;
  end;

// *********************************************************************//
// The Class CoPCsObject provides a Create and CreateRemote method to
// create instances of the default interface IPCsObject exposed by
// the CoClass PCsObject. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsObject = class
    class function Create: IPCsObject;
    class function CreateRemote(const MachineName: string): IPCsObject;
  end;

// *********************************************************************//
// The Class CoApp provides a Create and CreateRemote method to
// create instances of the default interface IApp exposed by
// the CoClass App. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoApp = class
    class function Create: IApp;
    class function CreateRemote(const MachineName: string): IApp;
  end;

// *********************************************************************//
// The Class CoPCsDatafield provides a Create and CreateRemote method to
// create instances of the default interface IPCsDatafield exposed by
// the CoClass PCsDatafield. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsDatafield = class
    class function Create: IPCsDatafield;
    class function CreateRemote(const MachineName: string): IPCsDatafield;
  end;

// *********************************************************************//
// The Class CoPCsDataLink provides a Create and CreateRemote method to
// create instances of the default interface IPCsDataLink exposed by
// the CoClass PCsDataLink. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsDataLink = class
    class function Create: IPCsDataLink;
    class function CreateRemote(const MachineName: string): IPCsDataLink;
  end;

// *********************************************************************//
// The Class CoPCsReferenceDesignation provides a Create and CreateRemote method to
// create instances of the default interface IPCsReferenceDesignation exposed by
// the CoClass PCsReferenceDesignation. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsReferenceDesignation = class
    class function Create: IPCsReferenceDesignation;
    class function CreateRemote(const MachineName: string): IPCsReferenceDesignation;
  end;

// *********************************************************************//
// The Class CoPCsSymbolFolderAliases provides a Create and CreateRemote method to
// create instances of the default interface IPCsSymbolFolderAliases exposed by
// the CoClass PCsSymbolFolderAliases. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsSymbolFolderAliases = class
    class function Create: IPCsSymbolFolderAliases;
    class function CreateRemote(const MachineName: string): IPCsSymbolFolderAliases;
  end;

// *********************************************************************//
// The Class CoPCsDialog provides a Create and CreateRemote method to
// create instances of the default interface IPCsDialog exposed by
// the CoClass PCsDialog. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsDialog = class
    class function Create: IPCsDialog;
    class function CreateRemote(const MachineName: string): IPCsDialog;
  end;

// *********************************************************************//
// The Class CoPCsFileDialog provides a Create and CreateRemote method to
// create instances of the default interface IPCsFileDialog exposed by
// the CoClass PCsFileDialog. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsFileDialog = class
    class function Create: IPCsFileDialog;
    class function CreateRemote(const MachineName: string): IPCsFileDialog;
  end;

// *********************************************************************//
// The Class CoPCsLineTexts provides a Create and CreateRemote method to
// create instances of the default interface IPCsLineTexts exposed by
// the CoClass PCsLineTexts. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsLineTexts = class
    class function Create: IPCsLineTexts;
    class function CreateRemote(const MachineName: string): IPCsLineTexts;
  end;

// *********************************************************************//
// The Class CoPCsExplorerWindow provides a Create and CreateRemote method to
// create instances of the default interface IPCsExplorerWindow exposed by
// the CoClass PCsExplorerWindow. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsExplorerWindow = class
    class function Create: IPCsExplorerWindow;
    class function CreateRemote(const MachineName: string): IPCsExplorerWindow;
  end;

// *********************************************************************//
// The Class CoPCsSubDrawing provides a Create and CreateRemote method to
// create instances of the default interface IPCsSubDrawing exposed by
// the CoClass PCsSubDrawing. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsSubDrawing = class
    class function Create: IPCsSubDrawing;
    class function CreateRemote(const MachineName: string): IPCsSubDrawing;
  end;

// *********************************************************************//
// The Class CoPCsListConnData provides a Create and CreateRemote method to
// create instances of the default interface IPCsListConnData exposed by
// the CoClass PCsListConnData. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsListConnData = class
    class function Create: IPCsListConnData;
    class function CreateRemote(const MachineName: string): IPCsListConnData;
  end;

// *********************************************************************//
// The Class CoPCsSelectStatus provides a Create and CreateRemote method to
// create instances of the default interface IPCsSelectStatus exposed by
// the CoClass PCsSelectStatus. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsSelectStatus = class
    class function Create: IPCsSelectStatus;
    class function CreateRemote(const MachineName: string): IPCsSelectStatus;
  end;

// *********************************************************************//
// The Class CoPCsSelectionArea provides a Create and CreateRemote method to
// create instances of the default interface IPCsSelectionArea exposed by
// the CoClass PCsSelectionArea. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsSelectionArea = class
    class function Create: IPCsSelectionArea;
    class function CreateRemote(const MachineName: string): IPCsSelectionArea;
  end;

// *********************************************************************//
// The Class CoPCsJoins provides a Create and CreateRemote method to
// create instances of the default interface IPCsJoins exposed by
// the CoClass PCsJoins. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsJoins = class
    class function Create: IPCsJoins;
    class function CreateRemote(const MachineName: string): IPCsJoins;
  end;

// *********************************************************************//
// The Class CoPCsJoin provides a Create and CreateRemote method to
// create instances of the default interface IPCsJoin exposed by
// the CoClass PCsJoin. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsJoin = class
    class function Create: IPCsJoin;
    class function CreateRemote(const MachineName: string): IPCsJoin;
  end;

// *********************************************************************//
// The Class CoPCsDrawingPreferences provides a Create and CreateRemote method to
// create instances of the default interface IPCsDrawingPreferences exposed by
// the CoClass PCsDrawingPreferences. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsDrawingPreferences = class
    class function Create: IPCsDrawingPreferences;
    class function CreateRemote(const MachineName: string): IPCsDrawingPreferences;
  end;

// *********************************************************************//
// The Class CoPCsReference provides a Create and CreateRemote method to
// create instances of the default interface IPCsReference exposed by
// the CoClass PCsReference. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsReference = class
    class function Create: IPCsReference;
    class function CreateRemote(const MachineName: string): IPCsReference;
  end;

// *********************************************************************//
// The Class CoPCsReferenceFrame provides a Create and CreateRemote method to
// create instances of the default interface IPCsReferenceFrame exposed by
// the CoClass PCsReferenceFrame. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsReferenceFrame = class
    class function Create: IPCsReferenceFrame;
    class function CreateRemote(const MachineName: string): IPCsReferenceFrame;
  end;

// *********************************************************************//
// The Class CoPCsDocument provides a Create and CreateRemote method to
// create instances of the default interface IPCsDocument exposed by
// the CoClass PCsDocument. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsDocument = class
    class function Create: IPCsDocument;
    class function CreateRemote(const MachineName: string): IPCsDocument;
  end;

// *********************************************************************//
// The Class CoPCsFigure provides a Create and CreateRemote method to
// create instances of the default interface IPCsFigure exposed by
// the CoClass PCsFigure. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsFigure = class
    class function Create: IPCsFigure;
    class function CreateRemote(const MachineName: string): IPCsFigure;
  end;

// *********************************************************************//
// The Class CoPCsList provides a Create and CreateRemote method to
// create instances of the default interface IPCsList exposed by
// the CoClass PCsList. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsList = class
    class function Create: IPCsList;
    class function CreateRemote(const MachineName: string): IPCsList;
  end;

// *********************************************************************//
// The Class CoPCsLibType provides a Create and CreateRemote method to
// create instances of the default interface IPCsLibType exposed by
// the CoClass PCsLibType. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsLibType = class
    class function Create: IPCsLibType;
    class function CreateRemote(const MachineName: string): IPCsLibType;
  end;

// *********************************************************************//
// The Class CoPCsTemplateData provides a Create and CreateRemote method to
// create instances of the default interface IPCsTemplateData exposed by
// the CoClass PCsTemplateData. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsTemplateData = class
    class function Create: IPCsTemplateData;
    class function CreateRemote(const MachineName: string): IPCsTemplateData;
  end;

// *********************************************************************//
// The Class CoPCsPages provides a Create and CreateRemote method to
// create instances of the default interface IPCsPages exposed by
// the CoClass PCsPages. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsPages = class
    class function Create: IPCsPages;
    class function CreateRemote(const MachineName: string): IPCsPages;
  end;

// *********************************************************************//
// The Class CoPCsPicture provides a Create and CreateRemote method to
// create instances of the default interface IPCsPicture exposed by
// the CoClass PCsPicture. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsPicture = class
    class function Create: IPCsPicture;
    class function CreateRemote(const MachineName: string): IPCsPicture;
  end;

// *********************************************************************//
// The Class CoPCsDatabaseSetup provides a Create and CreateRemote method to
// create instances of the default interface IPCsDatabaseSetup exposed by
// the CoClass PCsDatabaseSetup. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsDatabaseSetup = class
    class function Create: IPCsDatabaseSetup;
    class function CreateRemote(const MachineName: string): IPCsDatabaseSetup;
  end;

// *********************************************************************//
// The Class CoApplication provides a Create and CreateRemote method to
// create instances of the default interface IPCsApplication exposed by
// the CoClass Application. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoApplication = class
    class function Create: IPCsApplication;
    class function CreateRemote(const MachineName: string): IPCsApplication;
  end;

// *********************************************************************//
// The Class CoPCsDrawing provides a Create and CreateRemote method to
// create instances of the default interface IPCsDrawing exposed by
// the CoClass PCsDrawing. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsDrawing = class
    class function Create: IPCsDrawing;
    class function CreateRemote(const MachineName: string): IPCsDrawing;
  end;

// *********************************************************************//
// The Class CoPCsComponentDatabase provides a Create and CreateRemote method to
// create instances of the default interface IPCsComponentDatabase exposed by
// the CoClass PCsComponentDatabase. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsComponentDatabase = class
    class function Create: IPCsComponentDatabase;
    class function CreateRemote(const MachineName: string): IPCsComponentDatabase;
  end;

// *********************************************************************//
// The Class CoPCsLists provides a Create and CreateRemote method to
// create instances of the default interface IPCsLists exposed by
// the CoClass PCsLists. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsLists = class
    class function Create: IPCsLists;
    class function CreateRemote(const MachineName: string): IPCsLists;
  end;

// *********************************************************************//
// The Class CoPCsStatusObject provides a Create and CreateRemote method to
// create instances of the default interface IPCsStatusObject exposed by
// the CoClass PCsStatusObject. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsStatusObject = class
    class function Create: IPCsStatusObject;
    class function CreateRemote(const MachineName: string): IPCsStatusObject;
  end;

// *********************************************************************//
// The Class CoPCsConnection provides a Create and CreateRemote method to
// create instances of the default interface IPCsConnection exposed by
// the CoClass PCsConnection. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsConnection = class
    class function Create: IPCsConnection;
    class function CreateRemote(const MachineName: string): IPCsConnection;
  end;

// *********************************************************************//
// The Class CoPCsExternalObjects provides a Create and CreateRemote method to
// create instances of the default interface IPCsExternalObjects exposed by
// the CoClass PCsExternalObjects. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsExternalObjects = class
    class function Create: IPCsExternalObjects;
    class function CreateRemote(const MachineName: string): IPCsExternalObjects;
  end;

// *********************************************************************//
// The Class CoPCsExternalObject provides a Create and CreateRemote method to
// create instances of the default interface IPCsExternalObject exposed by
// the CoClass PCsExternalObject. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsExternalObject = class
    class function Create: IPCsExternalObject;
    class function CreateRemote(const MachineName: string): IPCsExternalObject;
  end;

// *********************************************************************//
// The Class CoPCsComponentName provides a Create and CreateRemote method to
// create instances of the default interface IPCsComponentName exposed by
// the CoClass PCsComponentName. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsComponentName = class
    class function Create: IPCsComponentName;
    class function CreateRemote(const MachineName: string): IPCsComponentName;
  end;

// *********************************************************************//
// The Class CoPCsSymbol provides a Create and CreateRemote method to
// create instances of the default interface IPCsSymbol exposed by
// the CoClass PCsSymbol. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsSymbol = class
    class function Create: IPCsSymbol;
    class function CreateRemote(const MachineName: string): IPCsSymbol;
  end;

// *********************************************************************//
// The Class CoPCsObjectRefID provides a Create and CreateRemote method to
// create instances of the default interface IPCsObjectRefID exposed by
// the CoClass PCsObjectRefID. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsObjectRefID = class
    class function Create: IPCsObjectRefID;
    class function CreateRemote(const MachineName: string): IPCsObjectRefID;
  end;

// *********************************************************************//
// The Class CoPCsReferenceDesignationList provides a Create and CreateRemote method to
// create instances of the default interface IPCsReferenceDesignationList exposed by
// the CoClass PCsReferenceDesignationList. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsReferenceDesignationList = class
    class function Create: IPCsReferenceDesignationList;
    class function CreateRemote(const MachineName: string): IPCsReferenceDesignationList;
  end;

// *********************************************************************//
// The Class CoPCsText provides a Create and CreateRemote method to
// create instances of the default interface IPCsText exposed by
// the CoClass PCsText. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsText = class
    class function Create: IPCsText;
    class function CreateRemote(const MachineName: string): IPCsText;
  end;

// *********************************************************************//
// The Class CoPCsArc provides a Create and CreateRemote method to
// create instances of the default interface IPCsArc exposed by
// the CoClass PCsArc. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsArc = class
    class function Create: IPCsArc;
    class function CreateRemote(const MachineName: string): IPCsArc;
  end;

// *********************************************************************//
// The Class CoPCsSymbols provides a Create and CreateRemote method to
// create instances of the default interface IPCsSymbols exposed by
// the CoClass PCsSymbols. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsSymbols = class
    class function Create: IPCsSymbols;
    class function CreateRemote(const MachineName: string): IPCsSymbols;
  end;

// *********************************************************************//
// The Class CoPCsComponentSymbols provides a Create and CreateRemote method to
// create instances of the default interface IPCsComponentSymbols exposed by
// the CoClass PCsComponentSymbols. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsComponentSymbols = class
    class function Create: IPCsComponentSymbols;
    class function CreateRemote(const MachineName: string): IPCsComponentSymbols;
  end;

// *********************************************************************//
// The Class CoPCsSymbolsEnumerator provides a Create and CreateRemote method to
// create instances of the default interface IPCsSymbolsEnumerator exposed by
// the CoClass PCsSymbolsEnumerator. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsSymbolsEnumerator = class
    class function Create: IPCsSymbolsEnumerator;
    class function CreateRemote(const MachineName: string): IPCsSymbolsEnumerator;
  end;

// *********************************************************************//
// The Class CoPCsLine provides a Create and CreateRemote method to
// create instances of the default interface IPCsLine exposed by
// the CoClass PCsLine. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsLine = class
    class function Create: IPCsLine;
    class function CreateRemote(const MachineName: string): IPCsLine;
  end;

// *********************************************************************//
// The Class CoPCsDatafields provides a Create and CreateRemote method to
// create instances of the default interface IPCsDatafields exposed by
// the CoClass PCsDatafields. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsDatafields = class
    class function Create: IPCsDatafields;
    class function CreateRemote(const MachineName: string): IPCsDatafields;
  end;

// *********************************************************************//
// The Class CoPCsLineDatafields provides a Create and CreateRemote method to
// create instances of the default interface IPCsLineDatafields exposed by
// the CoClass PCsLineDatafields. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsLineDatafields = class
    class function Create: IPCsLineDatafields;
    class function CreateRemote(const MachineName: string): IPCsLineDatafields;
  end;

// *********************************************************************//
// The Class CoPCsNameText provides a Create and CreateRemote method to
// create instances of the default interface IPCsNameText exposed by
// the CoClass PCsNameText. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsNameText = class
    class function Create: IPCsNameText;
    class function CreateRemote(const MachineName: string): IPCsNameText;
  end;

// *********************************************************************//
// The Class CoPCsRefIDFormat provides a Create and CreateRemote method to
// create instances of the default interface IPCsRefIDFormat exposed by
// the CoClass PCsRefIDFormat. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsRefIDFormat = class
    class function Create: IPCsRefIDFormat;
    class function CreateRemote(const MachineName: string): IPCsRefIDFormat;
  end;

// *********************************************************************//
// The Class CoPCsGlobal provides a Create and CreateRemote method to
// create instances of the default interface IPCsGlobal exposed by
// the CoClass PCsGlobal. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsGlobal = class
    class function Create: IPCsGlobal;
    class function CreateRemote(const MachineName: string): IPCsGlobal;
  end;

// *********************************************************************//
// The Class CoPCsReferenceDesignations provides a Create and CreateRemote method to
// create instances of the default interface IPCsReferenceDesignations exposed by
// the CoClass PCsReferenceDesignations. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsReferenceDesignations = class
    class function Create: IPCsReferenceDesignations;
    class function CreateRemote(const MachineName: string): IPCsReferenceDesignations;
  end;

// *********************************************************************//
// The Class CoPCsPage provides a Create and CreateRemote method to
// create instances of the default interface IPCsPage exposed by
// the CoClass PCsPage. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsPage = class
    class function Create: IPCsPage;
    class function CreateRemote(const MachineName: string): IPCsPage;
  end;

// *********************************************************************//
// The Class CoPCsComponent provides a Create and CreateRemote method to
// create instances of the default interface IPCsComponent exposed by
// the CoClass PCsComponent. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsComponent = class
    class function Create: IPCsComponent;
    class function CreateRemote(const MachineName: string): IPCsComponent;
  end;

// *********************************************************************//
// The Class CoPCsCompName provides a Create and CreateRemote method to
// create instances of the default interface IPCsCompName exposed by
// the CoClass PCsCompName. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoPCsCompName = class
    class function Create: IPCsCompName;
    class function CreateRemote(const MachineName: string): IPCsCompName;
  end;

implementation

uses System.Win.ComObj;

class function CoPCsPoints.Create: IPCsPoints;
begin
  Result := CreateComObject(CLASS_PCsPoints) as IPCsPoints;
end;

class function CoPCsPoints.CreateRemote(const MachineName: string): IPCsPoints;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsPoints) as IPCsPoints;
end;

class function CoPCsDocuments.Create: IPCsDocuments;
begin
  Result := CreateComObject(CLASS_PCsDocuments) as IPCsDocuments;
end;

class function CoPCsDocuments.CreateRemote(const MachineName: string): IPCsDocuments;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsDocuments) as IPCsDocuments;
end;

class function CoPCsApplicationPreferences.Create: IPCsApplicationPreferences;
begin
  Result := CreateComObject(CLASS_PCsApplicationPreferences) as IPCsApplicationPreferences;
end;

class function CoPCsApplicationPreferences.CreateRemote(const MachineName: string): IPCsApplicationPreferences;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsApplicationPreferences) as IPCsApplicationPreferences;
end;

class function CoPCsStatusPoint.Create: IPCsStatusPoint;
begin
  Result := CreateComObject(CLASS_PCsStatusPoint) as IPCsStatusPoint;
end;

class function CoPCsStatusPoint.CreateRemote(const MachineName: string): IPCsStatusPoint;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsStatusPoint) as IPCsStatusPoint;
end;

class function CoPCsLibTypes.Create: IPCsLibTypes;
begin
  Result := CreateComObject(CLASS_PCsLibTypes) as IPCsLibTypes;
end;

class function CoPCsLibTypes.CreateRemote(const MachineName: string): IPCsLibTypes;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsLibTypes) as IPCsLibTypes;
end;

class function CoPCsLines.Create: IPCsLines;
begin
  Result := CreateComObject(CLASS_PCsLines) as IPCsLines;
end;

class function CoPCsLines.CreateRemote(const MachineName: string): IPCsLines;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsLines) as IPCsLines;
end;

class function CoPCsArcs.Create: IPCsArcs;
begin
  Result := CreateComObject(CLASS_PCsArcs) as IPCsArcs;
end;

class function CoPCsArcs.CreateRemote(const MachineName: string): IPCsArcs;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsArcs) as IPCsArcs;
end;

class function CoPCsPageSetup.Create: IPCsPageSetup;
begin
  Result := CreateComObject(CLASS_PCsPageSetup) as IPCsPageSetup;
end;

class function CoPCsPageSetup.CreateRemote(const MachineName: string): IPCsPageSetup;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsPageSetup) as IPCsPageSetup;
end;

class function CoPCsPageReference.Create: IPCsPageReference;
begin
  Result := CreateComObject(CLASS_PCsPageReference) as IPCsPageReference;
end;

class function CoPCsPageReference.CreateRemote(const MachineName: string): IPCsPageReference;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsPageReference) as IPCsPageReference;
end;

class function CoPCsBasePageRef.Create: IPCsBasePageRef;
begin
  Result := CreateComObject(CLASS_PCsBasePageRef) as IPCsBasePageRef;
end;

class function CoPCsBasePageRef.CreateRemote(const MachineName: string): IPCsBasePageRef;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsBasePageRef) as IPCsBasePageRef;
end;

class function CoPCsTexts.Create: IPCsTexts;
begin
  Result := CreateComObject(CLASS_PCsTexts) as IPCsTexts;
end;

class function CoPCsTexts.CreateRemote(const MachineName: string): IPCsTexts;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsTexts) as IPCsTexts;
end;

class function CoPCsSymbolTexts.Create: IPCsSymbolTexts;
begin
  Result := CreateComObject(CLASS_PCsSymbolTexts) as IPCsSymbolTexts;
end;

class function CoPCsSymbolTexts.CreateRemote(const MachineName: string): IPCsSymbolTexts;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsSymbolTexts) as IPCsSymbolTexts;
end;

class function CoPCsSymbolMenu.Create: IPCsSymbolMenu;
begin
  Result := CreateComObject(CLASS_PCsSymbolMenu) as IPCsSymbolMenu;
end;

class function CoPCsSymbolMenu.CreateRemote(const MachineName: string): IPCsSymbolMenu;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsSymbolMenu) as IPCsSymbolMenu;
end;

class function CoPCsLibTypeConnections.Create: IPCsLibTypeConnections;
begin
  Result := CreateComObject(CLASS_PCsLibTypeConnections) as IPCsLibTypeConnections;
end;

class function CoPCsLibTypeConnections.CreateRemote(const MachineName: string): IPCsLibTypeConnections;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsLibTypeConnections) as IPCsLibTypeConnections;
end;

class function CoPCsConnections.Create: IPCsConnections;
begin
  Result := CreateComObject(CLASS_PCsConnections) as IPCsConnections;
end;

class function CoPCsConnections.CreateRemote(const MachineName: string): IPCsConnections;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsConnections) as IPCsConnections;
end;

class function CoPCsConnectionTexts.Create: IPCsConnectionTexts;
begin
  Result := CreateComObject(CLASS_PCsConnectionTexts) as IPCsConnectionTexts;
end;

class function CoPCsConnectionTexts.CreateRemote(const MachineName: string): IPCsConnectionTexts;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsConnectionTexts) as IPCsConnectionTexts;
end;

class function CoPCsTextFont.Create: IPCsTextFont;
begin
  Result := CreateComObject(CLASS_PCsTextFont) as IPCsTextFont;
end;

class function CoPCsTextFont.CreateRemote(const MachineName: string): IPCsTextFont;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsTextFont) as IPCsTextFont;
end;

class function CoPCsJoinPoint.Create: IPCsJoinPoint;
begin
  Result := CreateComObject(CLASS_PCsJoinPoint) as IPCsJoinPoint;
end;

class function CoPCsJoinPoint.CreateRemote(const MachineName: string): IPCsJoinPoint;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsJoinPoint) as IPCsJoinPoint;
end;

class function CoPCsLayers.Create: IPCsLayers;
begin
  Result := CreateComObject(CLASS_PCsLayers) as IPCsLayers;
end;

class function CoPCsLayers.CreateRemote(const MachineName: string): IPCsLayers;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsLayers) as IPCsLayers;
end;

class function CoPCsLayer.Create: IPCsLayer;
begin
  Result := CreateComObject(CLASS_PCsLayer) as IPCsLayer;
end;

class function CoPCsLayer.CreateRemote(const MachineName: string): IPCsLayer;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsLayer) as IPCsLayer;
end;

class function CoPCsObject.Create: IPCsObject;
begin
  Result := CreateComObject(CLASS_PCsObject) as IPCsObject;
end;

class function CoPCsObject.CreateRemote(const MachineName: string): IPCsObject;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsObject) as IPCsObject;
end;

class function CoApp.Create: IApp;
begin
  Result := CreateComObject(CLASS_App) as IApp;
end;

class function CoApp.CreateRemote(const MachineName: string): IApp;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_App) as IApp;
end;

class function CoPCsDatafield.Create: IPCsDatafield;
begin
  Result := CreateComObject(CLASS_PCsDatafield) as IPCsDatafield;
end;

class function CoPCsDatafield.CreateRemote(const MachineName: string): IPCsDatafield;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsDatafield) as IPCsDatafield;
end;

class function CoPCsDataLink.Create: IPCsDataLink;
begin
  Result := CreateComObject(CLASS_PCsDataLink) as IPCsDataLink;
end;

class function CoPCsDataLink.CreateRemote(const MachineName: string): IPCsDataLink;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsDataLink) as IPCsDataLink;
end;

class function CoPCsReferenceDesignation.Create: IPCsReferenceDesignation;
begin
  Result := CreateComObject(CLASS_PCsReferenceDesignation) as IPCsReferenceDesignation;
end;

class function CoPCsReferenceDesignation.CreateRemote(const MachineName: string): IPCsReferenceDesignation;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsReferenceDesignation) as IPCsReferenceDesignation;
end;

class function CoPCsSymbolFolderAliases.Create: IPCsSymbolFolderAliases;
begin
  Result := CreateComObject(CLASS_PCsSymbolFolderAliases) as IPCsSymbolFolderAliases;
end;

class function CoPCsSymbolFolderAliases.CreateRemote(const MachineName: string): IPCsSymbolFolderAliases;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsSymbolFolderAliases) as IPCsSymbolFolderAliases;
end;

class function CoPCsDialog.Create: IPCsDialog;
begin
  Result := CreateComObject(CLASS_PCsDialog) as IPCsDialog;
end;

class function CoPCsDialog.CreateRemote(const MachineName: string): IPCsDialog;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsDialog) as IPCsDialog;
end;

class function CoPCsFileDialog.Create: IPCsFileDialog;
begin
  Result := CreateComObject(CLASS_PCsFileDialog) as IPCsFileDialog;
end;

class function CoPCsFileDialog.CreateRemote(const MachineName: string): IPCsFileDialog;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsFileDialog) as IPCsFileDialog;
end;

class function CoPCsLineTexts.Create: IPCsLineTexts;
begin
  Result := CreateComObject(CLASS_PCsLineTexts) as IPCsLineTexts;
end;

class function CoPCsLineTexts.CreateRemote(const MachineName: string): IPCsLineTexts;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsLineTexts) as IPCsLineTexts;
end;

class function CoPCsExplorerWindow.Create: IPCsExplorerWindow;
begin
  Result := CreateComObject(CLASS_PCsExplorerWindow) as IPCsExplorerWindow;
end;

class function CoPCsExplorerWindow.CreateRemote(const MachineName: string): IPCsExplorerWindow;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsExplorerWindow) as IPCsExplorerWindow;
end;

class function CoPCsSubDrawing.Create: IPCsSubDrawing;
begin
  Result := CreateComObject(CLASS_PCsSubDrawing) as IPCsSubDrawing;
end;

class function CoPCsSubDrawing.CreateRemote(const MachineName: string): IPCsSubDrawing;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsSubDrawing) as IPCsSubDrawing;
end;

class function CoPCsListConnData.Create: IPCsListConnData;
begin
  Result := CreateComObject(CLASS_PCsListConnData) as IPCsListConnData;
end;

class function CoPCsListConnData.CreateRemote(const MachineName: string): IPCsListConnData;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsListConnData) as IPCsListConnData;
end;

class function CoPCsSelectStatus.Create: IPCsSelectStatus;
begin
  Result := CreateComObject(CLASS_PCsSelectStatus) as IPCsSelectStatus;
end;

class function CoPCsSelectStatus.CreateRemote(const MachineName: string): IPCsSelectStatus;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsSelectStatus) as IPCsSelectStatus;
end;

class function CoPCsSelectionArea.Create: IPCsSelectionArea;
begin
  Result := CreateComObject(CLASS_PCsSelectionArea) as IPCsSelectionArea;
end;

class function CoPCsSelectionArea.CreateRemote(const MachineName: string): IPCsSelectionArea;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsSelectionArea) as IPCsSelectionArea;
end;

class function CoPCsJoins.Create: IPCsJoins;
begin
  Result := CreateComObject(CLASS_PCsJoins) as IPCsJoins;
end;

class function CoPCsJoins.CreateRemote(const MachineName: string): IPCsJoins;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsJoins) as IPCsJoins;
end;

class function CoPCsJoin.Create: IPCsJoin;
begin
  Result := CreateComObject(CLASS_PCsJoin) as IPCsJoin;
end;

class function CoPCsJoin.CreateRemote(const MachineName: string): IPCsJoin;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsJoin) as IPCsJoin;
end;

class function CoPCsDrawingPreferences.Create: IPCsDrawingPreferences;
begin
  Result := CreateComObject(CLASS_PCsDrawingPreferences) as IPCsDrawingPreferences;
end;

class function CoPCsDrawingPreferences.CreateRemote(const MachineName: string): IPCsDrawingPreferences;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsDrawingPreferences) as IPCsDrawingPreferences;
end;

class function CoPCsReference.Create: IPCsReference;
begin
  Result := CreateComObject(CLASS_PCsReference) as IPCsReference;
end;

class function CoPCsReference.CreateRemote(const MachineName: string): IPCsReference;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsReference) as IPCsReference;
end;

class function CoPCsReferenceFrame.Create: IPCsReferenceFrame;
begin
  Result := CreateComObject(CLASS_PCsReferenceFrame) as IPCsReferenceFrame;
end;

class function CoPCsReferenceFrame.CreateRemote(const MachineName: string): IPCsReferenceFrame;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsReferenceFrame) as IPCsReferenceFrame;
end;

class function CoPCsDocument.Create: IPCsDocument;
begin
  Result := CreateComObject(CLASS_PCsDocument) as IPCsDocument;
end;

class function CoPCsDocument.CreateRemote(const MachineName: string): IPCsDocument;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsDocument) as IPCsDocument;
end;

class function CoPCsFigure.Create: IPCsFigure;
begin
  Result := CreateComObject(CLASS_PCsFigure) as IPCsFigure;
end;

class function CoPCsFigure.CreateRemote(const MachineName: string): IPCsFigure;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsFigure) as IPCsFigure;
end;

class function CoPCsList.Create: IPCsList;
begin
  Result := CreateComObject(CLASS_PCsList) as IPCsList;
end;

class function CoPCsList.CreateRemote(const MachineName: string): IPCsList;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsList) as IPCsList;
end;

class function CoPCsLibType.Create: IPCsLibType;
begin
  Result := CreateComObject(CLASS_PCsLibType) as IPCsLibType;
end;

class function CoPCsLibType.CreateRemote(const MachineName: string): IPCsLibType;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsLibType) as IPCsLibType;
end;

class function CoPCsTemplateData.Create: IPCsTemplateData;
begin
  Result := CreateComObject(CLASS_PCsTemplateData) as IPCsTemplateData;
end;

class function CoPCsTemplateData.CreateRemote(const MachineName: string): IPCsTemplateData;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsTemplateData) as IPCsTemplateData;
end;

class function CoPCsPages.Create: IPCsPages;
begin
  Result := CreateComObject(CLASS_PCsPages) as IPCsPages;
end;

class function CoPCsPages.CreateRemote(const MachineName: string): IPCsPages;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsPages) as IPCsPages;
end;

class function CoPCsPicture.Create: IPCsPicture;
begin
  Result := CreateComObject(CLASS_PCsPicture) as IPCsPicture;
end;

class function CoPCsPicture.CreateRemote(const MachineName: string): IPCsPicture;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsPicture) as IPCsPicture;
end;

class function CoPCsDatabaseSetup.Create: IPCsDatabaseSetup;
begin
  Result := CreateComObject(CLASS_PCsDatabaseSetup) as IPCsDatabaseSetup;
end;

class function CoPCsDatabaseSetup.CreateRemote(const MachineName: string): IPCsDatabaseSetup;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsDatabaseSetup) as IPCsDatabaseSetup;
end;

class function CoApplication.Create: IPCsApplication;
begin
  Result := CreateComObject(CLASS_Application) as IPCsApplication;
end;

class function CoApplication.CreateRemote(const MachineName: string): IPCsApplication;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Application) as IPCsApplication;
end;

class function CoPCsDrawing.Create: IPCsDrawing;
begin
  Result := CreateComObject(CLASS_PCsDrawing) as IPCsDrawing;
end;

class function CoPCsDrawing.CreateRemote(const MachineName: string): IPCsDrawing;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsDrawing) as IPCsDrawing;
end;

class function CoPCsComponentDatabase.Create: IPCsComponentDatabase;
begin
  Result := CreateComObject(CLASS_PCsComponentDatabase) as IPCsComponentDatabase;
end;

class function CoPCsComponentDatabase.CreateRemote(const MachineName: string): IPCsComponentDatabase;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsComponentDatabase) as IPCsComponentDatabase;
end;

class function CoPCsLists.Create: IPCsLists;
begin
  Result := CreateComObject(CLASS_PCsLists) as IPCsLists;
end;

class function CoPCsLists.CreateRemote(const MachineName: string): IPCsLists;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsLists) as IPCsLists;
end;

class function CoPCsStatusObject.Create: IPCsStatusObject;
begin
  Result := CreateComObject(CLASS_PCsStatusObject) as IPCsStatusObject;
end;

class function CoPCsStatusObject.CreateRemote(const MachineName: string): IPCsStatusObject;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsStatusObject) as IPCsStatusObject;
end;

class function CoPCsConnection.Create: IPCsConnection;
begin
  Result := CreateComObject(CLASS_PCsConnection) as IPCsConnection;
end;

class function CoPCsConnection.CreateRemote(const MachineName: string): IPCsConnection;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsConnection) as IPCsConnection;
end;

class function CoPCsExternalObjects.Create: IPCsExternalObjects;
begin
  Result := CreateComObject(CLASS_PCsExternalObjects) as IPCsExternalObjects;
end;

class function CoPCsExternalObjects.CreateRemote(const MachineName: string): IPCsExternalObjects;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsExternalObjects) as IPCsExternalObjects;
end;

class function CoPCsExternalObject.Create: IPCsExternalObject;
begin
  Result := CreateComObject(CLASS_PCsExternalObject) as IPCsExternalObject;
end;

class function CoPCsExternalObject.CreateRemote(const MachineName: string): IPCsExternalObject;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsExternalObject) as IPCsExternalObject;
end;

class function CoPCsComponentName.Create: IPCsComponentName;
begin
  Result := CreateComObject(CLASS_PCsComponentName) as IPCsComponentName;
end;

class function CoPCsComponentName.CreateRemote(const MachineName: string): IPCsComponentName;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsComponentName) as IPCsComponentName;
end;

class function CoPCsSymbol.Create: IPCsSymbol;
begin
  Result := CreateComObject(CLASS_PCsSymbol) as IPCsSymbol;
end;

class function CoPCsSymbol.CreateRemote(const MachineName: string): IPCsSymbol;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsSymbol) as IPCsSymbol;
end;

class function CoPCsObjectRefID.Create: IPCsObjectRefID;
begin
  Result := CreateComObject(CLASS_PCsObjectRefID) as IPCsObjectRefID;
end;

class function CoPCsObjectRefID.CreateRemote(const MachineName: string): IPCsObjectRefID;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsObjectRefID) as IPCsObjectRefID;
end;

class function CoPCsReferenceDesignationList.Create: IPCsReferenceDesignationList;
begin
  Result := CreateComObject(CLASS_PCsReferenceDesignationList) as IPCsReferenceDesignationList;
end;

class function CoPCsReferenceDesignationList.CreateRemote(const MachineName: string): IPCsReferenceDesignationList;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsReferenceDesignationList) as IPCsReferenceDesignationList;
end;

class function CoPCsText.Create: IPCsText;
begin
  Result := CreateComObject(CLASS_PCsText) as IPCsText;
end;

class function CoPCsText.CreateRemote(const MachineName: string): IPCsText;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsText) as IPCsText;
end;

class function CoPCsArc.Create: IPCsArc;
begin
  Result := CreateComObject(CLASS_PCsArc) as IPCsArc;
end;

class function CoPCsArc.CreateRemote(const MachineName: string): IPCsArc;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsArc) as IPCsArc;
end;

class function CoPCsSymbols.Create: IPCsSymbols;
begin
  Result := CreateComObject(CLASS_PCsSymbols) as IPCsSymbols;
end;

class function CoPCsSymbols.CreateRemote(const MachineName: string): IPCsSymbols;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsSymbols) as IPCsSymbols;
end;

class function CoPCsComponentSymbols.Create: IPCsComponentSymbols;
begin
  Result := CreateComObject(CLASS_PCsComponentSymbols) as IPCsComponentSymbols;
end;

class function CoPCsComponentSymbols.CreateRemote(const MachineName: string): IPCsComponentSymbols;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsComponentSymbols) as IPCsComponentSymbols;
end;

class function CoPCsSymbolsEnumerator.Create: IPCsSymbolsEnumerator;
begin
  Result := CreateComObject(CLASS_PCsSymbolsEnumerator) as IPCsSymbolsEnumerator;
end;

class function CoPCsSymbolsEnumerator.CreateRemote(const MachineName: string): IPCsSymbolsEnumerator;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsSymbolsEnumerator) as IPCsSymbolsEnumerator;
end;

class function CoPCsLine.Create: IPCsLine;
begin
  Result := CreateComObject(CLASS_PCsLine) as IPCsLine;
end;

class function CoPCsLine.CreateRemote(const MachineName: string): IPCsLine;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsLine) as IPCsLine;
end;

class function CoPCsDatafields.Create: IPCsDatafields;
begin
  Result := CreateComObject(CLASS_PCsDatafields) as IPCsDatafields;
end;

class function CoPCsDatafields.CreateRemote(const MachineName: string): IPCsDatafields;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsDatafields) as IPCsDatafields;
end;

class function CoPCsLineDatafields.Create: IPCsLineDatafields;
begin
  Result := CreateComObject(CLASS_PCsLineDatafields) as IPCsLineDatafields;
end;

class function CoPCsLineDatafields.CreateRemote(const MachineName: string): IPCsLineDatafields;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsLineDatafields) as IPCsLineDatafields;
end;

class function CoPCsNameText.Create: IPCsNameText;
begin
  Result := CreateComObject(CLASS_PCsNameText) as IPCsNameText;
end;

class function CoPCsNameText.CreateRemote(const MachineName: string): IPCsNameText;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsNameText) as IPCsNameText;
end;

class function CoPCsRefIDFormat.Create: IPCsRefIDFormat;
begin
  Result := CreateComObject(CLASS_PCsRefIDFormat) as IPCsRefIDFormat;
end;

class function CoPCsRefIDFormat.CreateRemote(const MachineName: string): IPCsRefIDFormat;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsRefIDFormat) as IPCsRefIDFormat;
end;

class function CoPCsGlobal.Create: IPCsGlobal;
begin
  Result := CreateComObject(CLASS_PCsGlobal) as IPCsGlobal;
end;

class function CoPCsGlobal.CreateRemote(const MachineName: string): IPCsGlobal;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsGlobal) as IPCsGlobal;
end;

class function CoPCsReferenceDesignations.Create: IPCsReferenceDesignations;
begin
  Result := CreateComObject(CLASS_PCsReferenceDesignations) as IPCsReferenceDesignations;
end;

class function CoPCsReferenceDesignations.CreateRemote(const MachineName: string): IPCsReferenceDesignations;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsReferenceDesignations) as IPCsReferenceDesignations;
end;

class function CoPCsPage.Create: IPCsPage;
begin
  Result := CreateComObject(CLASS_PCsPage) as IPCsPage;
end;

class function CoPCsPage.CreateRemote(const MachineName: string): IPCsPage;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsPage) as IPCsPage;
end;

class function CoPCsComponent.Create: IPCsComponent;
begin
  Result := CreateComObject(CLASS_PCsComponent) as IPCsComponent;
end;

class function CoPCsComponent.CreateRemote(const MachineName: string): IPCsComponent;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsComponent) as IPCsComponent;
end;

class function CoPCsCompName.Create: IPCsCompName;
begin
  Result := CreateComObject(CLASS_PCsCompName) as IPCsCompName;
end;

class function CoPCsCompName.CreateRemote(const MachineName: string): IPCsCompName;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PCsCompName) as IPCsCompName;
end;

end.

