using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Xml;
using System.Xml.Linq;

namespace OfficePicture.Core;

public static class OpenXmlImageExtractor
{
    private const string PackageNamespace = "http://schemas.microsoft.com/office/2006/xmlPackage";
    private const string OfficeRelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public static bool TryExtractWordImage(
        string flatOpenXml,
        int? shapeId,
        string? shapeName,
        out Image? image) =>
        TryExtractWordImage(flatOpenXml, shapeId, shapeName, null, out image);

    public static bool TryExtractWordImageByOrdinal(
        string flatOpenXml,
        int pictureOrdinal,
        out Image? image) =>
        TryExtractWordImage(flatOpenXml, null, null, pictureOrdinal, out image);

    public static bool TryExtractWordImageByOrdinalFromPackage(
        string packagePath,
        int pictureOrdinal,
        out Image? image)
        => TryExtractWordImageFromPackage(
            packagePath,
            pictureOrdinal,
            null,
            null,
            out image);

    public static bool TryExtractWordImageByIdentityFromPackage(
        string packagePath,
        string? anchorId,
        string? editId,
        out Image? image)
        => TryExtractWordImageFromPackage(
            packagePath,
            null,
            anchorId,
            editId,
            out image);

    private static bool TryExtractWordImageFromPackage(
        string packagePath,
        int? pictureOrdinal,
        string? anchorId,
        string? editId,
        out Image? image)
    {
        image = null;
        if ((!pictureOrdinal.HasValue || pictureOrdinal.Value < 1) &&
            string.IsNullOrWhiteSpace(anchorId) &&
            string.IsNullOrWhiteSpace(editId))
            return false;

        try
        {
            using var package = OpenPackage(packagePath);
            const string documentPart = "word/document.xml";
            var document = LoadXml(package, documentPart);
            if (document is null) return false;

            var pictures = document.Descendants()
                .Where(element => element.Name.LocalName == "inline");
            var picture = !string.IsNullOrWhiteSpace(anchorId) || !string.IsNullOrWhiteSpace(editId)
                ? pictures.FirstOrDefault(element => DrawingIdentityMatches(element, anchorId, editId))
                : pictures.Skip(pictureOrdinal!.Value - 1).FirstOrDefault();
            if (picture is null) return false;

            var imageParts = GetEmbeddedRelationshipIds(picture)
                .Select(id => ResolveRelationship(package, documentPart, id))
                .Where(path => path is not null)
                .Cast<string>();

            return TryChooseLargestImage(
                imageParts.Select(path => ReadEntryBytes(package, path)).Where(bytes => bytes is not null).Cast<byte[]>(),
                out image);
        }
        catch (Exception exception) when (
            exception is IOException ||
            exception is InvalidDataException ||
            exception is UnauthorizedAccessException ||
            exception is ArgumentException ||
            exception is XmlException)
        {
            return false;
        }
    }

    private static bool DrawingIdentityMatches(XElement element, string? anchorId, string? editId)
    {
        var elementAnchorId = element.Attributes()
            .FirstOrDefault(attribute => attribute.Name.LocalName == "anchorId")?.Value;
        var elementEditId = element.Attributes()
            .FirstOrDefault(attribute => attribute.Name.LocalName == "editId")?.Value;
        var compared = false;

        if (!string.IsNullOrWhiteSpace(anchorId) && !string.IsNullOrWhiteSpace(elementAnchorId))
        {
            compared = true;
            if (!string.Equals(elementAnchorId, anchorId, StringComparison.OrdinalIgnoreCase))
                return false;
        }

        if (!string.IsNullOrWhiteSpace(editId) && !string.IsNullOrWhiteSpace(elementEditId))
        {
            compared = true;
            if (!string.Equals(elementEditId, editId, StringComparison.OrdinalIgnoreCase))
                return false;
        }

        return compared;
    }

    private static bool TryExtractWordImage(
        string flatOpenXml,
        int? shapeId,
        string? shapeName,
        int? pictureOrdinal,
        out Image? image)
    {
        image = null;
        if (string.IsNullOrWhiteSpace(flatOpenXml)) return false;
        if (pictureOrdinal.HasValue && pictureOrdinal.Value < 1) return false;

        try
        {
            var package = XDocument.Parse(flatOpenXml, LoadOptions.PreserveWhitespace);
            var parts = GetFlatPackageParts(package);
            var candidatePartNames = GetWordImagePartNames(parts, shapeId, shapeName, pictureOrdinal);

            // InlineShape has no stable Office shape ID. Resolve the blip present in the
            // selected range XML instead of falling back to the first package media part.
            if (candidatePartNames.Count == 0 && (shapeId.HasValue || !string.IsNullOrEmpty(shapeName)))
                candidatePartNames = GetWordImagePartNames(parts, null, null, null);

            var candidates = candidatePartNames
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Select(name => parts.TryGetValue(NormalizePartName(name), out var part) ? part.Bytes : null)
                .Where(bytes => bytes is not null)
                .Cast<byte[]>();

            if (TryChooseLargestImage(candidates, out image)) return true;

            // Last-resort support for old Word picture formats that do not expose a blip.
            // Only use it when the Flat OPC package contains exactly one decodable image.
            var mediaParts = parts.Values
                .Where(part => part.ContentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase) && part.Bytes is not null)
                .Select(part => part.Bytes!)
                .ToList();
            if (mediaParts.Count != 1) return false;

            return TryChooseLargestImage(
                mediaParts,
                out image);
        }
        catch (Exception exception) when (
            exception is ArgumentException ||
            exception is FormatException ||
            exception is InvalidOperationException ||
            exception is XmlException)
        {
            return false;
        }
    }

    public static bool TryExtractPowerPointImage(
        string packagePath,
        int slideIndex,
        int shapeId,
        string shapeName,
        out Image? image)
    {
        image = null;
        try
        {
            using var package = OpenPackage(packagePath);
            var slidePart = GetPowerPointSlidePart(package, slideIndex);
            if (slidePart is null) return false;
            var slide = LoadXml(package, slidePart);
            if (slide is null) return false;

            var picture = FindPictureContainer(slide, shapeId, shapeName);
            if (picture is null) return false;

            var imageParts = GetEmbeddedRelationshipIds(picture)
                .Select(id => ResolveRelationship(package, slidePart, id))
                .Where(path => path is not null)
                .Cast<string>();

            return TryChooseLargestImage(
                imageParts.Select(path => ReadEntryBytes(package, path)).Where(bytes => bytes is not null).Cast<byte[]>(),
                out image);
        }
        catch (Exception exception) when (
            exception is IOException ||
            exception is InvalidDataException ||
            exception is UnauthorizedAccessException ||
            exception is ArgumentException ||
            exception is XmlException)
        {
            return false;
        }
    }

    public static bool TryExtractExcelImage(
        string packagePath,
        string worksheetName,
        int shapeId,
        string shapeName,
        out Image? image)
    {
        image = null;
        try
        {
            using var package = OpenPackage(packagePath);
            const string workbookPart = "xl/workbook.xml";
            var workbook = LoadXml(package, workbookPart);
            if (workbook is null) return false;

            var sheet = workbook.Descendants().FirstOrDefault(element =>
                element.Name.LocalName == "sheet" &&
                string.Equals((string?)element.Attribute("name"), worksheetName, StringComparison.Ordinal));
            var sheetRelationshipId = GetRelationshipAttribute(sheet, "id");
            if (sheetRelationshipId is null) return false;

            var worksheetPart = ResolveRelationship(package, workbookPart, sheetRelationshipId);
            if (worksheetPart is null) return false;
            var worksheet = LoadXml(package, worksheetPart);
            if (worksheet is null) return false;

            var drawingRelationshipId = worksheet.Descendants()
                .Where(element => element.Name.LocalName == "drawing")
                .Select(element => GetRelationshipAttribute(element, "id"))
                .FirstOrDefault(id => id is not null);
            if (drawingRelationshipId is null) return false;

            var drawingPart = ResolveRelationship(package, worksheetPart, drawingRelationshipId);
            if (drawingPart is null) return false;
            var drawing = LoadXml(package, drawingPart);
            if (drawing is null) return false;

            var picture = FindPictureContainer(drawing, shapeId, shapeName);
            if (picture is null) return false;

            var imageParts = GetEmbeddedRelationshipIds(picture)
                .Select(id => ResolveRelationship(package, drawingPart, id))
                .Where(path => path is not null)
                .Cast<string>();

            return TryChooseLargestImage(
                imageParts.Select(path => ReadEntryBytes(package, path)).Where(bytes => bytes is not null).Cast<byte[]>(),
                out image);
        }
        catch (Exception exception) when (
            exception is IOException ||
            exception is InvalidDataException ||
            exception is UnauthorizedAccessException ||
            exception is ArgumentException ||
            exception is XmlException)
        {
            return false;
        }
    }

    private static ZipArchive OpenPackage(string packagePath)
    {
        var stream = new FileStream(packagePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        try
        {
            return new ZipArchive(stream, ZipArchiveMode.Read, false);
        }
        catch
        {
            stream.Dispose();
            throw;
        }
    }

    private static List<string> GetWordImagePartNames(
        IReadOnlyDictionary<string, FlatPackagePart> parts,
        int? shapeId,
        string? shapeName,
        int? pictureOrdinal)
    {
        var result = new List<string>();
        var hasShapeIdentity = shapeId.HasValue || !string.IsNullOrEmpty(shapeName);

        foreach (var part in parts.Values.Where(item => item.Xml is not null))
        {
            IEnumerable<XElement> containers;
            if (pictureOrdinal.HasValue)
            {
                if (!string.Equals(part.Name, "/word/document.xml", StringComparison.OrdinalIgnoreCase))
                    continue;

                containers = part.Xml!.DescendantsAndSelf()
                    .Where(element => element.Name.LocalName == "inline")
                    .Skip(pictureOrdinal.Value - 1)
                    .Take(1);
            }
            else if (hasShapeIdentity)
            {
                containers = part.Xml!.DescendantsAndSelf()
                    .Where(element => IsShapePropertiesMatch(element, shapeId, shapeName))
                    .Select(properties => properties.AncestorsAndSelf().FirstOrDefault(IsPictureContainer))
                    .Where(container => container is not null)
                    .Cast<XElement>();
            }
            else
            {
                containers = part.Xml!.DescendantsAndSelf().Where(IsPictureContainer);
            }

            foreach (var relationshipId in containers.SelectMany(GetEmbeddedRelationshipIds).Distinct(StringComparer.Ordinal))
            {
                var target = ResolveFlatRelationship(parts, part.Name, relationshipId);
                if (target is not null) result.Add(target);
            }
        }

        return result.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
    }

    private static XElement? FindPictureContainer(XDocument document, int shapeId, string shapeName)
    {
        var properties = document.Descendants().FirstOrDefault(element =>
            IsShapePropertiesMatch(element, shapeId, shapeName));
        return properties?.AncestorsAndSelf().FirstOrDefault(IsPictureContainer);
    }

    private static string? GetPowerPointSlidePart(ZipArchive package, int slideIndex)
    {
        if (slideIndex < 1) return null;
        const string presentationPart = "ppt/presentation.xml";
        var presentation = LoadXml(package, presentationPart);
        if (presentation is null) return null;

        var slide = presentation.Descendants()
            .Where(element => element.Name.LocalName == "sldId")
            .Skip(slideIndex - 1)
            .FirstOrDefault();
        var relationshipId = GetRelationshipAttribute(slide, "id");
        return relationshipId is null
            ? null
            : ResolveRelationship(package, presentationPart, relationshipId);
    }

    private static bool IsShapePropertiesMatch(XElement element, int? shapeId, string? shapeName)
    {
        if (element.Name.LocalName != "cNvPr" && element.Name.LocalName != "docPr") return false;

        var idMatches = shapeId.HasValue &&
                        int.TryParse((string?)element.Attribute("id"), out var id) &&
                        id == shapeId.Value;
        var nameMatches = !string.IsNullOrEmpty(shapeName) &&
                          string.Equals((string?)element.Attribute("name"), shapeName, StringComparison.Ordinal);
        return idMatches || nameMatches;
    }

    private static bool IsPictureContainer(XElement element) =>
        element.Name.LocalName == "pic" ||
        element.Name.LocalName == "inline" ||
        element.Name.LocalName == "anchor";

    private static IEnumerable<string> GetEmbeddedRelationshipIds(XElement container) =>
        container.DescendantsAndSelf()
            .Attributes()
            .Where(attribute =>
                attribute.Name.LocalName == "embed" &&
                attribute.Name.NamespaceName == OfficeRelationshipsNamespace)
            .Select(attribute => attribute.Value)
            .Distinct(StringComparer.Ordinal);

    private static string? GetRelationshipAttribute(XElement? element, string localName) =>
        element?.Attributes().FirstOrDefault(attribute =>
            attribute.Name.LocalName == localName &&
            attribute.Name.NamespaceName == OfficeRelationshipsNamespace)?.Value;

    private static XDocument? LoadXml(ZipArchive package, string partPath)
    {
        var entry = package.GetEntry(NormalizeZipPath(partPath));
        if (entry is null) return null;
        using var stream = entry.Open();
        return XDocument.Load(stream, LoadOptions.None);
    }

    private static byte[]? ReadEntryBytes(ZipArchive package, string partPath)
    {
        var entry = package.GetEntry(NormalizeZipPath(partPath));
        if (entry is null) return null;
        using var source = entry.Open();
        using var destination = new MemoryStream();
        source.CopyTo(destination);
        return destination.ToArray();
    }

    private static string? ResolveRelationship(ZipArchive package, string sourcePart, string relationshipId)
    {
        var relationships = LoadXml(package, GetRelationshipPartPath(sourcePart));
        if (relationships is null) return null;

        var relationship = relationships.Descendants().FirstOrDefault(element =>
            element.Name.LocalName == "Relationship" &&
            string.Equals((string?)element.Attribute("Id"), relationshipId, StringComparison.Ordinal));
        if (relationship is null ||
            string.Equals((string?)relationship.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase))
            return null;

        var target = (string?)relationship.Attribute("Target");
        return target is null ? null : ResolvePartPath(sourcePart, target);
    }

    private static Dictionary<string, FlatPackagePart> GetFlatPackageParts(XDocument package)
    {
        XNamespace pkg = PackageNamespace;
        var result = new Dictionary<string, FlatPackagePart>(StringComparer.OrdinalIgnoreCase);

        foreach (var element in package.Descendants(pkg + "part"))
        {
            var name = NormalizePartName((string?)element.Attribute(pkg + "name") ?? string.Empty);
            if (name.Length <= 1) continue;

            var contentType = (string?)element.Attribute(pkg + "contentType") ?? string.Empty;
            var xmlRoot = element.Element(pkg + "xmlData")?.Elements().FirstOrDefault();
            var binaryText = element.Element(pkg + "binaryData")?.Value;
            byte[]? bytes = null;
            if (!string.IsNullOrWhiteSpace(binaryText))
            {
                try { bytes = Convert.FromBase64String(binaryText); }
                catch (FormatException) { }
            }

            result[name] = new FlatPackagePart(name, contentType, xmlRoot, bytes);
        }

        return result;
    }

    private static string? ResolveFlatRelationship(
        IReadOnlyDictionary<string, FlatPackagePart> parts,
        string sourcePart,
        string relationshipId)
    {
        var relationshipPartName = NormalizePartName(GetRelationshipPartPath(sourcePart));
        if (!parts.TryGetValue(relationshipPartName, out var relationshipPart) || relationshipPart.Xml is null)
            return null;

        var relationship = relationshipPart.Xml.DescendantsAndSelf().FirstOrDefault(element =>
            element.Name.LocalName == "Relationship" &&
            string.Equals((string?)element.Attribute("Id"), relationshipId, StringComparison.Ordinal));
        if (relationship is null ||
            string.Equals((string?)relationship.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase))
            return null;

        var target = (string?)relationship.Attribute("Target");
        return target is null ? null : NormalizePartName(ResolvePartPath(sourcePart, target));
    }

    private static string GetRelationshipPartPath(string sourcePart)
    {
        var normalized = NormalizeZipPath(sourcePart);
        var slash = normalized.LastIndexOf('/');
        var directory = slash < 0 ? string.Empty : normalized.Substring(0, slash + 1);
        var fileName = slash < 0 ? normalized : normalized.Substring(slash + 1);
        return $"{directory}_rels/{fileName}.rels";
    }

    private static string ResolvePartPath(string sourcePart, string target)
    {
        var sourceUri = new Uri("http://officepackage/" + NormalizeZipPath(sourcePart));
        var targetUri = new Uri(sourceUri, target.Replace('\\', '/'));
        return Uri.UnescapeDataString(targetUri.AbsolutePath.TrimStart('/'));
    }

    private static string NormalizeZipPath(string path) => path.Replace('\\', '/').TrimStart('/');
    private static string NormalizePartName(string path) => "/" + NormalizeZipPath(path);

    private static bool TryChooseLargestImage(IEnumerable<byte[]> candidates, out Image? image)
    {
        image = null;
        long largestArea = -1;

        foreach (var bytes in candidates)
        {
            var candidate = TryDecodeImage(bytes);
            if (candidate is null) continue;

            var area = (long)candidate.Width * candidate.Height;
            if (area > largestArea)
            {
                image?.Dispose();
                image = candidate;
                largestArea = area;
            }
            else
            {
                candidate.Dispose();
            }
        }

        return image is not null;
    }

    private static Image? TryDecodeImage(byte[] bytes)
    {
        try
        {
            using var stream = new MemoryStream(bytes, false);
            using var source = Image.FromStream(stream, true, true);
            return new Bitmap(source);
        }
        catch (Exception exception) when (
            exception is ArgumentException ||
            exception is ExternalException ||
            exception is OutOfMemoryException)
        {
            return null;
        }
    }

    private sealed class FlatPackagePart
    {
        public FlatPackagePart(string name, string contentType, XElement? xml, byte[]? bytes)
        {
            Name = name;
            ContentType = contentType;
            Xml = xml;
            Bytes = bytes;
        }

        public string Name { get; }
        public string ContentType { get; }
        public XElement? Xml { get; }
        public byte[]? Bytes { get; }
    }
}
