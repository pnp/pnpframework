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
            var serializers = GetWritersForCurrentContext(WriterScope.ProvisioningTemplate, a => a?.WriterSequence);

            // Invoke all the ProvisioningTemplate-level serializers
            InvokeWriters(template, writer, serializers);


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

        private IOrderedEnumerable<Type> GetWritersForCurrentContext(WriterScope scope,
            Func<TemplateSchemaWriterAttribute, Int32?> sortingSelector)
        {
            // Get all serializers to run in automated mode, ordered by sortingSelector
            var currentAssembly = this.GetType().Assembly;

            var writers = currentAssembly.GetTypes()
                // Get all the writers
                .Where(t => t.GetInterface(typeof(IPnPSchemaWriter).FullName) != null
                       && t.BaseType.Name == typeof(PnPBaseSchemaWriter<>).Name)
                // Order the writers by sequence
                .OrderBy(s =>
                {
                    var a = s.GetCustomAttributes<TemplateSchemaWriterAttribute>(false).FirstOrDefault();
                    return a.WriterSequence;
                }
                );
            
            return writers;
        }


        private static void InvokeWriters(ProvisioningTemplate template, TextWriter writer,
            IOrderedEnumerable<Type> writers)
        {
            foreach (var mdWriterType in writers)
            {
                // Get the first serializer only for each group (i.e. the most recent one for the current schema)
                var mdWriter = Activator.CreateInstance(mdWriterType) as IPnPSchemaWriter;
                if (mdWriter != null)
                {
                    mdWriter.Writer(template, writer);
                }
            }
        }
    }
}
