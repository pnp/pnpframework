using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using Newtonsoft.Json;
using PnP.Framework.Provisioning.Model;

namespace PnP.Framework.Provisioning.Providers.Markdown
{
    public class MarkdownPnPFormatter : ITemplateFormatterWithValidation
    {
        private TemplateProviderBase _provider;

        public void Initialize(TemplateProviderBase provider)
        {
            this._provider = provider;
        }

        public bool IsValid(Stream template)
        {
            return GetValidationResults(template).IsValid;
        }

        public ValidationResult GetValidationResults(System.IO.Stream template)
        {
            // We do not provide JSON validation capabilities
            return new ValidationResult { IsValid = true, Exceptions = null };
        }

        public System.IO.Stream ToFormattedTemplate(Model.ProvisioningTemplate template)
        {
            TextWriter writer = new StringWriter();
            // Get all ProvisioningTemplate-level serializers to run in automated mode, ordered by DeserializationSequence
            var serializers = GetSerializersForCurrentContext(WriterScope.ProvisioningTemplate, a => a?.WriterSequence);

            // Invoke all the ProvisioningTemplate-level serializers
            InvokeSerializers(template, writer, serializers);


            Byte[] markdownBytes = System.Text.Encoding.Unicode.GetBytes(writer.ToString());
            MemoryStream markdownStream = new MemoryStream(markdownBytes)
            {
                Position = 0
            };

            return (markdownStream);
        }

        public Model.ProvisioningTemplate ToProvisioningTemplate(System.IO.Stream template)
        {
            return (this.ToProvisioningTemplate(template, null));
        }

        public Model.ProvisioningTemplate ToProvisioningTemplate(System.IO.Stream template, string identifier)
        {
            StreamReader sr = new StreamReader(template, Encoding.Unicode);
            String jsonString = sr.ReadToEnd();
            Model.ProvisioningTemplate result = JsonConvert.DeserializeObject<Model.ProvisioningTemplate>(jsonString, new BasePermissionsConverter());
            return (result);
        }

        private IOrderedEnumerable<IGrouping<string, Type>> GetSerializersForCurrentContext(WriterScope scope,
            Func<TemplateSchemaWriterAttribute, Int32?> sortingSelector)
        {
            // Get all serializers to run in automated mode, ordered by sortingSelector
            var currentAssembly = this.GetType().Assembly;

            //PnP.Framework.Provisioning.Providers.Xml.XMLPnPSchemaVersion currentSchemaVersion = GetCurrentSchemaVersion();
            //TODO: finish off sorting

            var serializers = currentAssembly.GetTypes()
                // Get all the serializers
                .Where(t => t.GetInterface(typeof(IPnPSchemaWriter).FullName) != null)
                // Filter out those that are not targeting the current schema version or that are not in scope Template
                /*.Where(t =>
                {
                    var a = t.GetCustomAttributes<TemplateSchemaWriterAttribute>(false).FirstOrDefault();
                    return (a.Scope == scope);
                })*/
                // Order the remainings by supported schema version descendant, to get first the newest ones
                .OrderByDescending(s =>
                {

                    var a = s.Name;//.GetCustomAttributes<TemplateSchemaWriterAttribute>(false).FirstOrDefault();
                    return a;
                }
                )
                // Group those with the same target type (which is the first generic Type argument)
                .GroupBy(t => t.BaseType.GenericTypeArguments.FirstOrDefault()?.FullName)
            
                // Order the result by SerializationSequence
                .OrderBy(g =>
                {
                    /*var maxInGroup = g.OrderByDescending(s =>
                    {
                        var a = s.GetCustomAttributes<TemplateSchemaWriterAttribute>(false).FirstOrDefault();
                        return (a.MinimalSupportedSchemaVersion);
                    }
                    ).FirstOrDefault();*/
                    return g.Key;
                    //sortingSelector(maxInGroup.GetCustomAttributes<TemplateSchemaWriterAttribute>(false).FirstOrDefault());
                });
            
            return serializers;
        }


        private static void InvokeSerializers(ProvisioningTemplate template, TextWriter writer,
            IOrderedEnumerable<IGrouping<string, Type>> serializers)
        {
            foreach (var group in serializers)
            {
                // Get the first serializer only for each group (i.e. the most recent one for the current schema)
                var serializerType = group.FirstOrDefault();
                if (serializerType != null)
                {
                    // Create an instance of the serializer
                    var serializer = Activator.CreateInstance(serializerType) as IPnPSchemaWriter;
                    if (serializer != null)
                    {
                        serializer.Writer(template, writer);
                    }
                }
            }
        }
    }
}
